"""
goyo_extracted（NAIST誤用コーパス）と ORD_batch1_25.tsv の
語順誤り(ORD)データを比較する統計分析スクリプト
"""
import sys, csv, re, statistics
import openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. データ読み込み
# ============================================================

def load_goyo_ord(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = str(row[3]).lower() if row[3] else ''
        if t.startswith('ord'):
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


def load_artificial_ord(tsv_path):
    records = []
    with open(tsv_path, encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter='\t')
        next(reader, None)  # ヘッダースキップ
        for row in reader:
            if len(row) < 6:
                continue
            t = str(row[1]).lower() if row[1] else ''
            if t == 'ord':
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

# A型で使う格助詞・複合格助詞（長い方から先に照合）
_PARTICLES = sorted([
    'にとって', 'において', 'によって', 'を通じて', 'に関して',
    'に向けて', 'をめぐって', 'に比べて', 'に従って', 'に沿って',
    'について', 'として', 'より', 'から', 'まで', 'ほど', 'で', 'に', 'へ',
], key=len, reverse=True)

# C型で使う副詞リスト
_ADVERBS = ['みんな', 'どちらも', 'ほぼ', 'ほとんど', 'だいたい',
            'すっかり', 'やっと', 'ちょうど', 'まだ', 'もう']


def classify_ord(err_tok, cor_tok):
    # D型: 日付（例：15日の3月）または 時刻（例：30分の10時）
    if re.search(r'\d+[日分]の\d+[月時]', err_tok):
        return 'D型: 日付・時刻の語順（英語式）'

    # A型: 格助詞が名詞より前に出る
    for p in _PARTICLES:
        if err_tok.startswith(p):
            return 'A型: 格助詞の前置'

    # C型: 副詞が主語より前に出る
    for adv in _ADVERBS:
        if err_tok.startswith(adv):
            return 'C型: 副詞・量詞の位置ずれ'

    # B型: 名詞句の逆転（N1のN2 → N2のN1）
    # cor_tokも「の」を含む場合のみ真のB型（例：先端の時代↔時代の先端）
    # err_tokだけに「の」があり cor_tokに「の」がない場合はE型（形容詞後置など）
    if 'の' in err_tok and 'の' in cor_tok:
        return 'B型: 名詞句の語順逆転'

    return 'E型: その他語順誤り'


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
            st = classify_ord(r['err_tok'], r['cor_tok'])
            counts[st] = counts.get(st, 0) + 1
        return counts

    g_counts = count_subtypes(goyo_recs)
    a_counts = count_subtypes(art_recs)

    all_types = ['A型: 格助詞の前置', 'B型: 名詞句の語順逆転',
                 'C型: 副詞・量詞の位置ずれ', 'D型: 日付・時刻の語順（英語式）',
                 'E型: その他語順誤り']
    # goyo に出現した未知サブタイプも追加
    for st in sorted(set(g_counts) | set(a_counts)):
        if st not in all_types:
            all_types.append(st)

    g_total = sum(g_counts.values()) or 1
    a_total = sum(a_counts.values()) or 1

    print(f"\n  {'サブタイプ':<36}  {'goyo':>18}   {'人工':>12}   差")
    print(f"  {'-'*36}  {'-'*18}   {'-'*12}   {'-'*6}")
    for st in all_types:
        g_n   = g_counts.get(st, 0)
        a_n   = a_counts.get(st, 0)
        g_pct = g_n / g_total * 100
        a_pct = a_n / a_total * 100
        diff  = a_pct - g_pct
        sign  = '+' if diff >= 0 else ''
        print(f"  {st:<36}  {g_n:4d}件({g_pct:5.1f}%)   {a_n:3d}件({a_pct:5.1f}%)   {sign}{diff:.1f}")


# ============================================================
# 5. メイン
# ============================================================

def main():
    base      = Path(__file__).parent.parent
    goyo_path = base / 'goyo_extracted.xlsx'
    tsv_path  = base / 'tsv' / 'ORD_batch1_25.tsv'

    print("データ読み込み中...")
    goyo_recs = load_goyo_ord(goyo_path)
    art_recs  = load_artificial_ord(tsv_path)

    if not goyo_recs:
        print("WARNING: goyo_extracted に ORD型データが見つかりません（比較スキップ）")
    if not art_recs:
        print("ERROR: ORD_batch1_25.tsv に ORD型データが見つかりません")
        return

    art_err_len = [len(r['err']) for r in art_recs]
    art_cor_len = [len(r['cor']) for r in art_recs]
    as_e = calc_stats(art_err_len)
    as_c = calc_stats(art_cor_len)

    # ---- 人工データ単体 ----
    print(f"\n{SEP}")
    print(f" 人工データ（ORD_batch1_25.tsv）: 語順誤り(ORD)型  {as_e['n']}件")
    print(SEP)
    print_stats(as_e, '誤文')
    print_stats(as_c, '正文')

    if goyo_recs:
        goyo_err_len = [len(r['err']) for r in goyo_recs]
        goyo_cor_len = [len(r['cor']) for r in goyo_recs]
        gs_e = calc_stats(goyo_err_len)
        gs_c = calc_stats(goyo_cor_len)

        # ---- goyo ----
        print(f"\n{SEP}")
        print(f" NAIST誤用コーパス（goyo_extracted）: ORD型  {gs_e['n']}件")
        print(SEP)
        print_stats(gs_e, '誤文')
        print_stats(gs_c, '正文')

        # ---- 差分 ----
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

        # ---- ヒストグラム ----
        print(f"\n{SEP}")
        print(" 文字数ヒストグラム（10文字幅）")
        print(SEP)
        print_hist(gs_e, 'goyo 誤文')
        print_hist(as_e, '人工 誤文')
        print_hist(gs_c, 'goyo 正文')
        print_hist(as_c, '人工 正文')

        # ---- サブタイプ比較 ----
        print(f"\n{SEP}")
        print(" サブタイプ比較（goyo vs 人工データ）")
        print(SEP)
        print_subtype_compare(goyo_recs, art_recs)

    else:
        # goyo に ORD がない場合は人工データのみ
        print(f"\n{SEP}")
        print(" 文字数ヒストグラム（10文字幅）")
        print(SEP)
        print_hist(as_e, '人工 誤文')
        print_hist(as_c, '人工 正文')

    # ---- サブタイプ分布（人工データ）----
    print(f"\n{SEP}")
    print(" サブタイプ分布（人工データ 100件）")
    print(SEP)
    counts = {}
    nos    = {}
    for r in art_recs:
        st = classify_ord(r['err_tok'], r['cor_tok'])
        counts[st] = counts.get(st, 0) + 1
        nos.setdefault(st, []).append(r['no'])

    total = sum(counts.values())
    for st, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        sample = ','.join(nos[st][:6])
        ellip  = '...' if len(nos[st]) > 6 else ''
        print(f"  {st:<38}: {cnt:3d}件 ({cnt/total*100:4.1f}%)  No.{sample}{ellip}")


if __name__ == '__main__':
    main()
