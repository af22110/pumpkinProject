import numpy as np
import pandas as pd
from scipy.stats import ks_2samp, wasserstein_distance
from scipy.spatial.distance import cdist
from sklearn.preprocessing import StandardScaler
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent / "NAIST誤用コーパス" / "Distribution comparison"

FEATURE_COLS = [
    "op_insert", "op_delete", "op_replace",
    "change_len",
    "ratio_hiragana", "ratio_katakana", "ratio_kanji", "ratio_other",
    "len_diff",
]

# --- データ読み込み ---
art = pd.read_excel(DATA_DIR / "art_features.xlsx")[FEATURE_COLS].dropna()
goyo = pd.read_excel(DATA_DIR / "goyo_features.xlsx")[FEATURE_COLS].dropna()

print(f"art_features: {len(art)} 件")
print(f"goyo_features: {len(goyo)} 件")

art_arr = art.values
goyo_arr = goyo.values
n_art = len(art_arr)

# ベースライン分割（goyo A: 1700件, goyo B: 残り）
rng_bl = np.random.default_rng(42)
idx = rng_bl.permutation(len(goyo_arr))
goyo_a = goyo_arr[idx[:n_art]]
goyo_b = goyo_arr[idx[n_art:]]

# ==========================================================
# ① KS検定・Wasserstein距離（標準化前）
# ==========================================================
print("\n=== ①各特徴量ごとのKS検定・Wasserstein距離 ===")
header = f"{'特徴量':<16}| {'本番 KS D':>11} | {'本番 W':>11} | {'BL KS D':>9} | {'BL W':>9}"
print(header)
print("-" * 67)

for i, col in enumerate(FEATURE_COLS):
    x_art  = art_arr[:, i]
    x_goyo = goyo_arr[:, i]
    x_ba   = goyo_a[:, i]
    x_bb   = goyo_b[:, i]

    ks_real, _ = ks_2samp(x_art, x_goyo)
    w_real      = wasserstein_distance(x_art, x_goyo)
    ks_base, _  = ks_2samp(x_ba, x_bb)
    w_base      = wasserstein_distance(x_ba, x_bb)

    print(f"{col:<16}| {ks_real:>11.4f} | {w_real:>11.4f} | {ks_base:>9.4f} | {w_base:>9.4f}")

# ==========================================================
# ② Energy Distance（標準化後）
# ==========================================================
print("\n=== ②Energy Distance（9次元同時、標準化後） ===")

GOYO_SAMPLE = 3000

# art + goyo全体でfit
scaler = StandardScaler()
combined_scaled = scaler.fit_transform(np.vstack([art_arr, goyo_arr]))
art_scaled  = combined_scaled[:n_art]
goyo_scaled = combined_scaled[n_art:]

# goyoから3000件サンプリング（本番）
rng_ed = np.random.default_rng(42)
goyo_s = goyo_scaled[rng_ed.choice(len(goyo_scaled), size=GOYO_SAMPLE, replace=False)]

print(f"[注] goyo件数が大きいため、goyo側をランダムサンプリング（{GOYO_SAMPLE}件）して近似計算。\n")

def mean_cross_dist(A, B):
    return cdist(A, B, metric="euclidean").mean()

def mean_self_dist(A):
    d = cdist(A, A, metric="euclidean")
    n = len(A)
    return d.sum() / (n * (n - 1))

def energy_distance(X, Y):
    return 2 * mean_cross_dist(X, Y) - mean_self_dist(X) - mean_self_dist(Y)

ed_real = energy_distance(art_scaled, goyo_s)

# ベースライン（同じscalerで変換済みのgoyo_a / goyo_b）
goyo_a_scaled = scaler.transform(goyo_a)
goyo_b_scaled = scaler.transform(goyo_b)
goyo_b_s = goyo_b_scaled[rng_ed.choice(len(goyo_b_scaled), size=GOYO_SAMPLE, replace=False)]

ed_base = energy_distance(goyo_a_scaled, goyo_b_s)

print(f"本番（人工データ vs goyo）          ：ED = {ed_real:.6f}")
print(f"ベースライン（goyo A vs goyo B）    ：ED = {ed_base:.6f}")
