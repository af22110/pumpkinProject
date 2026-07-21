# -*- coding: utf-8 -*-
"""
人工データと誤用コーパスの分布比較
正誤ペアの差分から特徴量を抽出してベクトル化し、PCAで2次元に圧縮して可視化する。
"""
import difflib
import re
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import rcParams
from sklearn.decomposition import PCA

# 日本語フォント設定（文字化け対策）
rcParams["font.family"] = "MS Gothic"

BASE_DIR = Path(__file__).resolve().parent.parent
OUT_DIR = BASE_DIR / "Distribution comparison"
OUT_DIR.mkdir(exist_ok=True)

ART_PATH = BASE_DIR / "artificial data2.xlsx"
GOYO_PATH = BASE_DIR / "goyo_extracted.xlsx"

HIRAGANA = (0x3040, 0x309F)
KATAKANA = (0x30A0, 0x30FF)
KANJI = (0x4E00, 0x9FFF)


def char_type_ratios(text: str) -> np.ndarray:
    """ひらがな・カタカナ・漢字・その他の比率（4次元）を返す。"""
    counts = np.zeros(4, dtype=float)
    if len(text) == 0:
        return counts
    for ch in text:
        code = ord(ch)
        if HIRAGANA[0] <= code <= HIRAGANA[1]:
            counts[0] += 1
        elif KATAKANA[0] <= code <= KATAKANA[1]:
            counts[1] += 1
        elif KANJI[0] <= code <= KANJI[1]:
            counts[2] += 1
        else:
            counts[3] += 1
    return counts / len(text)


def extract_features(correct: str, wrong: str) -> np.ndarray | None:
    """正解文と誤り文から9次元の特徴量ベクトルを抽出する。

    差分が見つからない（完全一致）場合は None を返す。
    """
    matcher = difflib.SequenceMatcher(a=correct, b=wrong, autojunk=False)
    opcodes = matcher.get_opcodes()

    first_diff = None
    for tag, i1, i2, j1, j2 in opcodes:
        if tag != "equal":
            first_diff = (tag, i1, i2, j1, j2)
            break

    if first_diff is None:
        return None

    tag, i1, i2, j1, j2 = first_diff

    # 1. 編集操作タイプ one-hot（insert / delete / replace）
    op_onehot = np.array(
        [1.0 if tag == "insert" else 0.0,
         1.0 if tag == "delete" else 0.0,
         1.0 if tag == "replace" else 0.0]
    )

    correct_span = correct[i1:i2]
    wrong_span = wrong[j1:j2]

    # 2. 変更箇所の文字数
    change_len = float(max(len(correct_span), len(wrong_span)))

    # 3. 変更箇所の文字種比率（delete は正解側、それ以外は誤り側）
    target_span = correct_span if tag == "delete" else wrong_span
    ratios = char_type_ratios(target_span)

    # 4. 正誤の文長差
    len_diff = float(len(wrong) - len(correct))

    return np.concatenate([op_onehot, [change_len], ratios, [len_diff]])


def build_vectors(df: pd.DataFrame, correct_col: str, wrong_col: str) -> np.ndarray:
    vectors = []
    for correct, wrong in zip(df[correct_col], df[wrong_col]):
        vec = extract_features(str(correct), str(wrong))
        if vec is not None:
            vectors.append(vec)
    return np.array(vectors)


def preprocess(df: pd.DataFrame, correct_col: str, wrong_col: str) -> pd.DataFrame:
    df = df.dropna(subset=[correct_col, wrong_col])
    df = df[df[correct_col] != df[wrong_col]]
    return df


FEATURE_COLUMNS = [
    "op_insert", "op_delete", "op_replace",
    "change_len",
    "ratio_hiragana", "ratio_katakana", "ratio_kanji", "ratio_other",
    "len_diff",
]


def main():
    art_df = pd.read_excel(ART_PATH)
    goyo_df = pd.read_excel(GOYO_PATH)

    art_correct_col, art_wrong_col = "正解文", "誤り文"
    goyo_correct_col, goyo_wrong_col = "正解文（全誤り修正済）", "誤り文（着目誤り1つのみ）"

    art_df = preprocess(art_df, art_correct_col, art_wrong_col)
    goyo_df = preprocess(goyo_df, goyo_correct_col, goyo_wrong_col)

    art_vectors = build_vectors(art_df, art_correct_col, art_wrong_col)
    goyo_vectors = build_vectors(goyo_df, goyo_correct_col, goyo_wrong_col)

    print(f"人工データ: {len(art_df)}件 -> 特徴量抽出 {len(art_vectors)}件")
    print(f"goyoデータ: {len(goyo_df)}件 -> 特徴量抽出 {len(goyo_vectors)}件")

    # PCA は結合した全体に対して fit する
    combined = np.vstack([art_vectors, goyo_vectors])
    pca = PCA(n_components=2)
    combined_2d = pca.fit_transform(combined)

    art_2d = combined_2d[: len(art_vectors)]
    goyo_2d = combined_2d[len(art_vectors):]

    # 散布図
    fig, ax = plt.subplots(figsize=(9, 7))
    ax.scatter(
        art_2d[:, 0], art_2d[:, 1],
        alpha=0.4, s=15, color="#1f77b4",
        label=f"人工データ (n={len(art_2d)})",
    )
    ax.scatter(
        goyo_2d[:, 0], goyo_2d[:, 1],
        alpha=0.4, s=15, color="#d62728",
        label=f"誤用コーパス (n={len(goyo_2d)})",
    )
    ax.set_title("人工データと誤用コーパスの分布比較（PCA）")
    ax.set_xlabel(f"PC1 (寄与率 {pca.explained_variance_ratio_[0]:.1%})")
    ax.set_ylabel(f"PC2 (寄与率 {pca.explained_variance_ratio_[1]:.1%})")
    ax.legend()
    fig.tight_layout()

    png_path = OUT_DIR / "distribution_comparison.png"
    fig.savefig(png_path, dpi=200)
    plt.close(fig)

    # 特徴量行列を Excel として保存
    art_features_df = pd.DataFrame(art_vectors, columns=FEATURE_COLUMNS)
    goyo_features_df = pd.DataFrame(goyo_vectors, columns=FEATURE_COLUMNS)
    art_features_df.to_excel(OUT_DIR / "art_features.xlsx", index=False)
    goyo_features_df.to_excel(OUT_DIR / "goyo_features.xlsx", index=False)

    print(f"散布図を保存しました: {png_path}")
    print(f"特徴量ファイルを保存しました: {OUT_DIR / 'art_features.xlsx'}, {OUT_DIR / 'goyo_features.xlsx'}")


if __name__ == "__main__":
    main()
