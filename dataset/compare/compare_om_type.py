"""
goyo_extracted（NAIST誤用コーパス）と om_batch1_50.tsv の
不足誤り(OM)データを比較する統計分析スクリプト
"""

import sys
import csv
import openpyxl
import statistics
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')


# ============================================================
# 1. データ読み込み
# ============================================================

def load_goyo_om(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = str(row[3]).lower() if row[3] else ''
        if t.startswith('om'):
            goyo = row[6]
            sei  = row[7]
            if goyo is not None and sei is not None:
                records.append({'goyo': str(goyo), 'sei': str(sei)})
    wb.close()
    return records

def load_artificial_om(tsv_path):
    records = []
    with open(tsv_path, encoding='utf-8') as f:
        reader = csv.reader(f, delimiter='\t')
        for row in reader:
            if len(row) < 6:
                continue
            t = str(row[1]).lower() if row[1] else ''
            if t.startswith('om'):
                goyo = row[4]
                sei  = row[5]
                if goyo and sei:
                    records.append({'goyo': goyo, 'sei': sei})
    return records


# ============================================================
# 2. 統計計算
# ============================================================

def calc_stats(lengths):
    n  = len(lengths)
    sl = sorted(lengths)
    return {
        'n':      n,
        'mean':   statistics.mean(lengths),
        'median': statistics.median(lengths),
        'var_s':  statistics.variance(lengths),
        'std_s':  statistics.stdev(lengths),
        'min':    min(lengths),
        'max':    max(lengths),
        'range':  max(lengths) - min(lengths),
        'q1':     sl[int(n * 0.25)],
        'q3':     sl[int(n * 0.75)],
        'iqr':    sl[int(n * 0.75)] - sl[int(n * 0.25)],
        'buckets': _buckets(lengths),
    }

def _buckets(lengths):
    b = {}
    for v in lengths:
        k = (v // 10) * 10
        b[k] = b.get(k, 0) + 1
    return b


# ============================================================
# 3. 表示ユーティリティ
# ============================================================

SEP = '=' * 60

def print_stats(s, label):
    print(f"\n▼ {label}")
    print(f"  件数          : {s['n']}")
    print(f"  平均文字数    : {s['mean']:.2f}")
    print(f"  中央値        : {s['median']:.1f}")
    print(f"  標本分散      : {s['var_s']:.2f}")
    print(f"  標本標準偏差  : {s['std_s']:.2f}")
    print(f"  最小 / 最大   : {s['min']} / {s['max']}")
    print(f"  レンジ        : {s['range']}")
    print(f"  Q1 / Q3 / IQR : {s['q1']} / {s['q3']} / {s['iqr']}")

def print_diff(g, a, key, fmt='.2f'):
    gv, av = g[key], a[key]
    diff = av - gv
    sign = '+' if diff >= 0 else ''
    print(f"  {key:<14}: goyo={gv:{fmt}}  art={av:{fmt}}  差={sign}{diff:{fmt}}")

def print_hist(s, label):
    print(f"\n  【文字数ヒストグラム（10文字幅）: {label}】")
    for b in sorted(s['buckets']):
        bar = '■' * (s['buckets'][b] * 30 // s['n'])
        print(f"  {b:3d}〜{b+9:3d}文字 : {s['buckets'][b]:4d}件 {bar}")


# ============================================================
# 4. メイン
# ============================================================

def main():
    base      = Path(__file__).parent.parent
    goyo_path = base / 'goyo_extracted.xlsx'
    tsv_path  = base / 'om_batch1_50.tsv'

    print("データ読み込み中...")
    goyo_recs = load_goyo_om(goyo_path)
    art_recs  = load_artificial_om(tsv_path)

    if not goyo_recs:
        print("ERROR: goyo_extractedにOMデータが見つかりません")
        return
    if not art_recs:
        print("ERROR: om_batch1_50.tsvにOMデータが見つかりません")
        return

    goyo_err_len = [len(r['goyo']) for r in goyo_recs]
    goyo_cor_len = [len(r['sei'])  for r in goyo_recs]
    art_err_len  = [len(r['goyo']) for r in art_recs]
    art_cor_len  = [len(r['sei'])  for r in art_recs]

    gs_g = calc_stats(goyo_err_len)
    gs_s = calc_stats(goyo_cor_len)
    as_g = calc_stats(art_err_len)
    as_s = calc_stats(art_cor_len)

    print(f"\n{SEP}")
    print(" NAIST誤用コーパス（goyo_extracted）: 不足誤り(OM)型")
    print(SEP)
    print_stats(gs_g, '誤文')
    print_stats(gs_s, '正文')

    print(f"\n{SEP}")
    print(" 人工データ（om_batch1_50.tsv）: 不足誤り(OM)型")
    print(SEP)
    print_stats(as_g, '誤文')
    print_stats(as_s, '正文')

    print(f"\n{SEP}")
    print(" 差分比較（人工データ − goyo）")
    print(SEP)
    print("\n▼ 誤文")
    for key, fmt in [('mean','.2f'),('median','.1f'),('var_s','.2f'),
                     ('std_s','.2f'),('min','d'),('max','d'),('range','d'),('iqr','d')]:
        print_diff(gs_g, as_g, key, fmt)
    print("\n▼ 正文")
    for key, fmt in [('mean','.2f'),('median','.1f'),('var_s','.2f'),
                     ('std_s','.2f'),('min','d'),('max','d'),('range','d'),('iqr','d')]:
        print_diff(gs_s, as_s, key, fmt)

    print(f"\n{SEP}")
    print(" 文字数ヒストグラム")
    print(SEP)
    print_hist(gs_g, 'goyo誤文')
    print_hist(as_g, 'art誤文')
    print_hist(gs_s, 'goyo正文')
    print_hist(as_s, 'art正文')


if __name__ == '__main__':
    main()
