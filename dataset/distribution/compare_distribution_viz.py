# -*- coding: utf-8 -*-
"""
人工データと誤用コーパスの分布比較（可視化改善版）
誤用コーパスの件数が人工データを圧倒しているため、単純な重ね散布図では
人工データの点が埋没する。KDE密度・左右パネル・周辺分布の3種類の図で補う。

前提: compare_distribution.py を先に実行し、
      Distribution comparison/art_features.xlsx, goyo_features.xlsx が
      生成済みであること。
"""
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib import rcParams
from sklearn.decomposition import PCA
from sklearn.preprocessing import StandardScaler

rcParams["font.family"] = "MS Gothic"

BASE_DIR = Path(__file__).resolve().parent.parent
OUT_DIR = BASE_DIR / "Distribution comparison"

ART_FEATURES_PATH = OUT_DIR / "art_features.xlsx"
GOYO_FEATURES_PATH = OUT_DIR / "goyo_features.xlsx"


def compute_pca():
    """特徴量を標準化してからPCAで2次元に圧縮する。"""
    art_features = pd.read_excel(ART_FEATURES_PATH).to_numpy()
    goyo_features = pd.read_excel(GOYO_FEATURES_PATH).to_numpy()

    combined = np.vstack([art_features, goyo_features])
    combined_scaled = StandardScaler().fit_transform(combined)

    pca = PCA(n_components=2)
    combined_2d = pca.fit_transform(combined_scaled)

    art_2d = combined_2d[: len(art_features)]
    goyo_2d = combined_2d[len(art_features):]
    return art_2d, goyo_2d, pca


def axis_limits(art_2d: np.ndarray, goyo_2d: np.ndarray, margin: float = 0.15,
                 pct: float = 1.0):
    """外れ値に引っ張られないよう、パーセンタイルベースで表示範囲を決める。"""
    combined = np.vstack([art_2d, goyo_2d])
    x_min, x_max = np.percentile(combined[:, 0], [pct, 100 - pct])
    y_min, y_max = np.percentile(combined[:, 1], [pct, 100 - pct])
    x_pad = (x_max - x_min) * margin
    y_pad = (y_max - y_min) * margin
    return (x_min - x_pad, x_max + x_pad), (y_min - y_pad, y_max + y_pad)


def plot_overlay_density(art_2d, goyo_2d, pca, xlim, ylim):
    """改善1: goyoはKDE等高線、人工データは前面に散布点。"""
    fig, ax = plt.subplots(figsize=(9, 7))

    sns.kdeplot(
        x=goyo_2d[:, 0], y=goyo_2d[:, 1],
        fill=True, cmap="Reds", thresh=0.05, levels=10, alpha=0.7,
        clip=(xlim, ylim), ax=ax, zorder=1,
    )
    sns.kdeplot(
        x=goyo_2d[:, 0], y=goyo_2d[:, 1],
        color="firebrick", levels=6, linewidths=0.6,
        clip=(xlim, ylim), ax=ax, zorder=2,
    )
    ax.scatter(
        art_2d[:, 0], art_2d[:, 1],
        s=35, color="#1565C0", edgecolors="white", linewidths=0.3,
        alpha=0.85, zorder=3,
        label=f"人工データ (n={len(art_2d)})",
    )

    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_title("人工データと誤用コーパスの分布比較（密度＋散布, PCA）")
    ax.set_xlabel(f"PC1 (寄与率 {pca.explained_variance_ratio_[0]:.1%})")
    ax.set_ylabel(f"PC2 (寄与率 {pca.explained_variance_ratio_[1]:.1%})")

    # 凡例（goyoの密度パッチを手動で追加）
    from matplotlib.patches import Patch
    handles = [
        Patch(facecolor="firebrick", alpha=0.5, label=f"誤用コーパス 密度 (n={len(goyo_2d)})"),
        plt.Line2D(
            [], [], marker="o", linestyle="", color="#1565C0",
            markeredgecolor="white", markersize=8,
            label=f"人工データ (n={len(art_2d)})",
        ),
    ]
    ax.legend(handles=handles, loc="best")

    fig.tight_layout()
    out_path = OUT_DIR / "pca_overlay_density.png"
    fig.savefig(out_path, dpi=200)
    plt.close(fig)
    return out_path


def plot_side_by_side(art_2d, goyo_2d, pca, xlim, ylim):
    """改善2: goyo(KDE等高線) と 人工データ(散布) を左右パネルで比較。"""
    fig, axes = plt.subplots(1, 2, figsize=(15, 7), sharex=True, sharey=True)

    ax_goyo, ax_art = axes

    sns.kdeplot(
        x=goyo_2d[:, 0], y=goyo_2d[:, 1],
        fill=True, cmap="Reds", thresh=0.02, levels=12,
        clip=(xlim, ylim), ax=ax_goyo,
    )
    ax_goyo.set_title(f"誤用コーパス (n={len(goyo_2d)})")
    ax_goyo.set_xlabel(f"PC1 (寄与率 {pca.explained_variance_ratio_[0]:.1%})")
    ax_goyo.set_ylabel(f"PC2 (寄与率 {pca.explained_variance_ratio_[1]:.1%})")

    ax_art.scatter(
        art_2d[:, 0], art_2d[:, 1],
        s=20, color="#1565C0", alpha=0.6,
    )
    sns.kdeplot(
        x=art_2d[:, 0], y=art_2d[:, 1],
        color="#0D47A1", levels=6, linewidths=0.8,
        clip=(xlim, ylim), ax=ax_art,
    )
    ax_art.set_title(f"人工データ (n={len(art_2d)})")
    ax_art.set_xlabel(f"PC1 (寄与率 {pca.explained_variance_ratio_[0]:.1%})")

    for ax in axes:
        ax.set_xlim(*xlim)
        ax.set_ylim(*ylim)

    fig.suptitle("人工データと誤用コーパスの分布比較（左右パネル, PCA）")
    fig.tight_layout()
    out_path = OUT_DIR / "pca_side_by_side.png"
    fig.savefig(out_path, dpi=200)
    plt.close(fig)
    return out_path


def plot_marginal(art_2d, goyo_2d, pca, xlim, ylim):
    """改善3: jointplotで散布図＋PC1/PC2周辺分布を重ねて表示。"""
    df = pd.concat([
        pd.DataFrame({"PC1": art_2d[:, 0], "PC2": art_2d[:, 1],
                       "データ種別": f"人工データ (n={len(art_2d)})"}),
        pd.DataFrame({"PC1": goyo_2d[:, 0], "PC2": goyo_2d[:, 1],
                       "データ種別": f"誤用コーパス (n={len(goyo_2d)})"}),
    ], ignore_index=True)

    art_label = f"人工データ (n={len(art_2d)})"
    goyo_label = f"誤用コーパス (n={len(goyo_2d)})"
    palette = {
        art_label: "#1565C0",
        goyo_label: "#d62728",
    }

    # hue_order の先頭から順に描画されるため、goyo を先に描き、
    # 件数の少ない art を最後（＝前面）に描画する
    g = sns.jointplot(
        data=df, x="PC1", y="PC2", hue="データ種別",
        hue_order=[goyo_label, art_label],
        palette=palette,
        kind="scatter",
        alpha=0.5, s=20,
        marginal_kws=dict(fill=True, common_norm=False, alpha=0.4),
        height=8,
    )
    g.figure.suptitle("人工データと誤用コーパスの分布比較（周辺分布付き, PCA）", y=1.02)
    g.ax_joint.set_xlabel(f"PC1 (寄与率 {pca.explained_variance_ratio_[0]:.1%})")
    g.ax_joint.set_ylabel(f"PC2 (寄与率 {pca.explained_variance_ratio_[1]:.1%})")
    g.ax_joint.set_xlim(*xlim)
    g.ax_joint.set_ylim(*ylim)

    out_path = OUT_DIR / "pca_marginal.png"
    g.figure.savefig(out_path, dpi=200, bbox_inches="tight")
    plt.close(g.figure)
    return out_path


def main():
    art_2d, goyo_2d, pca = compute_pca()
    xlim, ylim = axis_limits(art_2d, goyo_2d)

    p1 = plot_overlay_density(art_2d, goyo_2d, pca, xlim, ylim)
    p2 = plot_side_by_side(art_2d, goyo_2d, pca, xlim, ylim)
    p3 = plot_marginal(art_2d, goyo_2d, pca, xlim, ylim)

    print(f"保存しました: {p1}")
    print(f"保存しました: {p2}")
    print(f"保存しました: {p3}")


if __name__ == "__main__":
    main()
