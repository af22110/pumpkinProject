"""
goyo_extracted（NAIST誤用コーパス）と DEM_100.tsv の
指示詞誤り(DEM)データを比較する統計分析スクリプト
"""
import sys, csv, statistics
import openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. データ読み込み
# ============================================================

def load_goyo_dem(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = str(row[3]).lower() if row[3] else ''
        if t.startswith('dem'):
            err_text = row[6]
            cor_text = row[7]
            err_tok  = str(row[4]) if row[4] else ''
            cor_tok  = str(row[5]) if row[5] else ''
            if err_text is not None and cor_text is not None:
                records.append({
                    'err':     str(err_text),
                    'cor':     str(cor_text),
                    'err_tok': err_tok,
                    'cor_tok': cor_tok,
                    'type':    t,
                })
    wb.close()
    return records


def load_artificial_dem(tsv_path):
    records = []
    with open(tsv_path, encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter='\t')
        next(reader, None)  # ヘッダースキップ
        for row in reader:
            if len(row) < 6:
                continue
            t = str(row[1]).lower() if row[1] else ''
            if t == 'dem':
                records.append({
                    'no':      row[0],
                    'err_tok': row[2],
                    'cor_tok': row[3],
                    'err':     row[4],
                    'cor':     row[5],
                })
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
# 3. サブタイプ分類
# ============================================================

# 長い表現を先に照合（部分一致誤りを防ぐため降順ソート）
_KONO = sorted([
    'こんなに', 'そんなに', 'あんなに',
    'こんな', 'そんな', 'あんな',
    'こういった', 'そういった', 'ああいった',  # 「こう」部分一致誤りを防ぐ
    'こういう', 'そういう', 'ああいう',
    'この', 'その', 'あの', 'どの',
], key=len, reverse=True)

_KORE = sorted([
    'これら', 'それら', 'あれら',
    'こちら', 'そちら', 'あちら',
    'これ', 'それ', 'あれ',
], key=len, reverse=True)

_KOKO = ['ここ', 'そこ', 'あそこ']

_KOU  = sorted(['こう', 'そう', 'ああ'], key=len, reverse=True)


def _contains_any(text, words):
    """text の中にいずれかの単語が含まれるか（部分一致）"""
    for w in words:
        if w in text:
            return True
    return False


def classify_dem(err_tok, cor_tok):
    et = str(err_tok) if err_tok else ''
    ct = str(cor_tok) if cor_tok else ''
    combined = et + ct

    if _contains_any(combined, _KONO):
        return 'A型: この/その/あの系'
    if _contains_any(combined, _KORE):
        return 'B型: これ/それ/あれ系'
    if _contains_any(combined, _KOKO):
        return 'C型: ここ/そこ/あそこ系'
    if _contains_any(combined, _KOU):
        return 'D型: こう/そう/ああ系'
    return 'E型: その他'


# ============================================================
# 4. 表示ユーティリティ
# ============================================================

SEP = '=' * 62

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
    diff   = av - gv
    sign   = '+' if diff >= 0 else ''
    ok = ''
    if key == 'mean':
        ok = '  ✓' if abs(diff) <= 3  else '  ✗'
    elif key == 'var_s':
        ok = '  ✓' if abs(diff) <= 30 else '  ✗'
    print(f"  {key:<14}: goyo={gv:{fmt}}  art={av:{fmt}}  差={sign}{diff:{fmt}}{ok}")


def print_hist(s, label):
    print(f"\n  【{label}】")
    for b in sorted(s['buckets']):
        bar = '■' * (s['buckets'][b] * 30 // s['n'])
        pct = s['buckets'][b] / s['n'] * 100
        print(f"  {b:3d}〜{b+9:3d}文字 : {s['buckets'][b]:4d}件 ({pct:4.1f}%) {bar}")


def print_subtype_compare(goyo_recs, art_recs):
    def count_subtypes(recs):
        counts = {}
        for r in recs:
            st = classify_dem(r['err_tok'], r['cor_tok'])
            counts[st] = counts.get(st, 0) + 1
        return counts

    g_counts = count_subtypes(goyo_recs)
    a_counts = count_subtypes(art_recs)

    all_types = [
        'A型: この/その/あの系',
        'B型: これ/それ/あれ系',
        'C型: ここ/そこ/あそこ系',
        'D型: こう/そう/ああ系',
        'E型: その他',
    ]
    for st in sorted(set(g_counts) | set(a_counts)):
        if st not in all_types:
            all_types.append(st)

    g_total = sum(g_counts.values()) or 1
    a_total = sum(a_counts.values()) or 1

    print(f"\n  {'サブタイプ':<30}  {'goyo':>18}   {'人工':>12}   差")
    print(f"  {'-'*30}  {'-'*18}   {'-'*12}   {'-'*6}")
    for st in all_types:
        g_n   = g_counts.get(st, 0)
        a_n   = a_counts.get(st, 0)
        g_pct = g_n / g_total * 100
        a_pct = a_n / a_total * 100
        diff  = a_pct - g_pct
        sign  = '+' if diff >= 0 else ''
        print(f"  {st:<30}  {g_n:4d}件({g_pct:5.1f}%)   {a_n:3d}件({a_pct:5.1f}%)   {sign}{diff:.1f}")


# ============================================================
# 5. メイン
# ============================================================

def main():
    base      = Path(__file__).parent.parent
    goyo_path = base / 'goyo_extracted.xlsx'
    tsv_path  = base / 'tsv' / 'DEM_100.tsv'

    print("データ読み込み中...")
    goyo_recs = load_goyo_dem(goyo_path)
    art_recs  = load_artificial_dem(tsv_path)

    if not goyo_recs:
        print("WARNING: goyo_extracted に DEM型データが見つかりません")
    if not art_recs:
        print("ERROR: DEM_100.tsv に DEM型データが見つかりません")
        return

    art_err_len = [len(r['err']) for r in art_recs]
    art_cor_len = [len(r['cor']) for r in art_recs]
    as_e = calc_stats(art_err_len)
    as_c = calc_stats(art_cor_len)

    print(f"\n{SEP}")
    print(f" 人工データ（DEM_100.tsv）: 指示詞誤り(DEM)型  {as_e['n']}件")
    print(SEP)
    print_stats(as_e, '誤文')
    print_stats(as_c, '正文')

    if goyo_recs:
        goyo_err_len = [len(r['err']) for r in goyo_recs]
        goyo_cor_len = [len(r['cor']) for r in goyo_recs]
        gs_e = calc_stats(goyo_err_len)
        gs_c = calc_stats(goyo_cor_len)

        print(f"\n{SEP}")
        print(f" NAIST誤用コーパス（goyo_extracted）: DEM型  {gs_e['n']}件")
        print(SEP)
        print_stats(gs_e, '誤文')
        print_stats(gs_c, '正文')

        print(f"\n{SEP}")
        print(" 差分比較（人工データ − goyo）")
        print(" ※ 合格基準: mean差 ±3以内、var_s差 ±30以内")
        print(SEP)
        print("\n▼ 誤文")
        for key, fmt in [('mean','.2f'),('median','.1f'),('var_s','.2f'),
                         ('std_s','.2f'),('min','d'),('max','d'),('range','d'),('iqr','d')]:
            print_diff(gs_e, as_e, key, fmt)
        print("\n▼ 正文")
        for key, fmt in [('mean','.2f'),('median','.1f'),('var_s','.2f'),
                         ('std_s','.2f'),('min','d'),('max','d'),('range','d'),('iqr','d')]:
            print_diff(gs_c, as_c, key, fmt)

        print(f"\n{SEP}")
        print(" 文字数ヒストグラム（10文字幅）")
        print(SEP)
        print_hist(gs_e, 'goyo 誤文')
        print_hist(as_e, '人工 誤文')
        print_hist(gs_c, 'goyo 正文')
        print_hist(as_c, '人工 正文')

        print(f"\n{SEP}")
        print(" サブタイプ比較（goyo vs 人工データ）")
        print(SEP)
        print_subtype_compare(goyo_recs, art_recs)

    print(f"\n{SEP}")
    print(" サブタイプ分布（人工データ 100件）")
    print(SEP)
    counts = {}
    nos    = {}
    for r in art_recs:
        st = classify_dem(r['err_tok'], r['cor_tok'])
        counts[st] = counts.get(st, 0) + 1
        nos.setdefault(st, []).append(r['no'])

    total = sum(counts.values())
    for st, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        sample = ','.join(nos[st][:6])
        ellip  = '...' if len(nos[st]) > 6 else ''
        print(f"  {st:<32}: {cnt:3d}件 ({cnt/total*100:4.1f}%)  No.{sample}{ellip}")


if __name__ == '__main__':
    main()
