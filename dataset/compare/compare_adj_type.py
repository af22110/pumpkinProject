"""
goyo_extracted（NAIST誤用コーパス）と adj_complete_100.tsv の
形容詞誤り(ADJ)データを比較する統計分析スクリプト
"""
import sys, csv, statistics
import openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. データ読み込み
# ============================================================

def load_goyo_adj(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = str(row[3]).lower() if row[3] else ''
        if t.startswith('adj'):
            err_text = row[6]
            cor_text = row[7]
            err_tok  = str(row[4]) if row[4] else ''
            cor_tok  = str(row[5]) if row[5] else ''
            if err_text is not None and cor_text is not None:
                records.append({
                    'err': str(err_text), 'cor': str(cor_text),
                    'err_tok': err_tok, 'cor_tok': cor_tok,
                    'type': t,
                })
    wb.close()
    return records


def load_artificial_adj(tsv_path):
    records = []
    with open(tsv_path, encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter='\t')
        for row in reader:
            if len(row) < 6:
                continue
            t = str(row[1]).lower() if row[1] else ''
            if t == 'adj':
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


# な形容詞語幹リスト（これらの語幹＋い形式はな形容詞の誤活用）
_NA_ADJ_STEMS = {
    '静か', '便利', 'きれい', '元気', '有名', '親切', '大変', '簡単',
    '素直', '危険', '熱心', '正直', '自由', '新鮮', '特別', '立派',
    '不便', '複雑', '丈夫', '賑やか', '地味', '重要', '必要', '上手',
    '得意', '苦手', '大切', '大事', '豊か', '穏やか', '华やか', '華やか',
}

def _is_na_adj_stem(tok):
    """トークンがな形容詞語幹+い の形式かどうかを判定"""
    for stem in _NA_ADJ_STEMS:
        if tok.endswith(stem + 'い') or tok.endswith(stem + 'い'):
            return True
        # 「〜くて」「〜くない」など語幹+く系も判定
        if stem + 'くて' in tok or stem + 'く' in tok:
            return True
    return False


def classify_adj(err_tok, cor_tok):
    e, c = err_tok, cor_tok

    # い形容詞 + でした（過去丁寧の誤り）
    if e.endswith('でした') and (c.endswith('かったです') or c.endswith('かった')):
        return 'い形容詞+でした（過去活用誤り）'

    # い形容詞 + だった（過去普通の誤り）
    if e.endswith('いだった') and c.endswith('かった'):
        return 'い形容詞+だった（過去活用誤り）'

    # い形容詞 + だ（断定）
    if e.endswith('いだ') and not c.endswith('いだ'):
        return 'い形容詞+だ（断定余剰）'

    # い形容詞 + です + と （丁寧形余剰）
    if e.endswith('いです') and not c.endswith('いです'):
        return 'い形容詞+です（丁寧形余剰）'

    # な形容詞 → い形容詞化活用（語幹リストで判定）
    # くて → で の修正（例：便利くて → 便利で）
    if e.endswith('くて') and c.endswith('で') and _is_na_adj_stem(e):
        return 'な形容詞のい形容詞化（〜くて）'
    # い → な/だ 修正（例：静かい → 静かな）— な形容詞語幹のもののみ
    if (c.endswith('な') or c.endswith('だ')) and _is_na_adj_stem(e):
        return 'な形容詞のい形容詞化（〜い）'

    # 比較・程度表現（いい重複 / いく）
    if 'いい' in e and e.count('い') > c.count('い'):
        return '修飾・比較表現の活用誤り（い重複）'
    if e.endswith('いく') and c.endswith('く'):
        return '修飾・比較表現の活用誤り（いく）'
    if e.endswith('いくない') or e.endswith('いくなかった'):
        return '修飾・比較表現の活用誤り（いくない）'

    # 〜いそう（様態・伝聞）
    if e.endswith('いそう') and c.endswith('そう'):
        return 'い形容詞+いそう（様態活用誤り）'

    # 感情形容詞の表現者誤り（〜がる未使用）
    if 'がる' in c or 'がって' in c or 'がっている' in c:
        return '感情形容詞の表現者誤り（〜がる未使用）'

    # 類義形容詞の意味的誤選択（上記に当てはまらない場合）
    return '類義形容詞の意味的誤選択'


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
            st = classify_adj(r['err_tok'], r['cor_tok'])
            counts[st] = counts.get(st, 0) + 1
        return counts

    g_counts = count_subtypes(goyo_recs)
    a_counts = count_subtypes(art_recs)

    all_types = sorted(
        set(g_counts) | set(a_counts),
        key=lambda x: -a_counts.get(x, 0)
    )

    g_total = sum(g_counts.values())
    a_total = sum(a_counts.values())

    print(f"\n  {'サブタイプ':<32}  {'goyo':>20}   {'人工':>14}   差")
    print(f"  {'-'*32}  {'-'*20}   {'-'*14}   {'-'*6}")
    for st in all_types:
        g_n   = g_counts.get(st, 0)
        a_n   = a_counts.get(st, 0)
        g_pct = g_n / g_total * 100 if g_total > 0 else 0
        a_pct = a_n / a_total * 100
        diff  = a_pct - g_pct
        sign  = '+' if diff >= 0 else ''
        print(f"  {st:<32}  {g_n:5d}件({g_pct:5.1f}%)   {a_n:3d}件({a_pct:5.1f}%)   {sign}{diff:.1f}")


# ============================================================
# 5. メイン
# ============================================================

def main():
    base      = Path(__file__).parent.parent
    goyo_path = base / 'goyo_extracted.xlsx'
    tsv_path  = base / 'tsv' / 'adj_complete_100.tsv'

    print("データ読み込み中...")
    goyo_recs = load_goyo_adj(goyo_path)
    art_recs  = load_artificial_adj(tsv_path)

    if not goyo_recs:
        print("WARNING: goyo_extracted に ADJ型データが見つかりません（比較スキップ）")
    if not art_recs:
        print("ERROR: adj_complete_100.tsv に ADJ型データが見つかりません")
        return

    art_err_len = [len(r['err']) for r in art_recs]
    art_cor_len = [len(r['cor']) for r in art_recs]
    as_g = calc_stats(art_err_len)
    as_s = calc_stats(art_cor_len)

    # ---- 人工データ単体 ----
    print(f"\n{SEP}")
    print(f" 人工データ（adj_complete_100.tsv）: 形容詞誤り(ADJ)型  {as_g['n']}件")
    print(SEP)
    print_stats(as_g, '誤文')
    print_stats(as_s, '正文')

    if goyo_recs:
        goyo_err_len = [len(r['err']) for r in goyo_recs]
        goyo_cor_len = [len(r['cor']) for r in goyo_recs]
        gs_g = calc_stats(goyo_err_len)
        gs_s = calc_stats(goyo_cor_len)

        # ---- goyo ----
        print(f"\n{SEP}")
        print(f" NAIST誤用コーパス（goyo_extracted）: ADJ型  {gs_g['n']}件")
        print(SEP)
        print_stats(gs_g, '誤文')
        print_stats(gs_s, '正文')

        # ---- 差分 ----
        print(f"\n{SEP}")
        print(" 差分比較（人工データ − goyo）")
        print(" ※ 合格基準: mean差 ±3以内、var_s差 ±30以内")
        print(SEP)
        print("\n▼ 誤文")
        for key, fmt in [('mean','.2f'),('median','.1f'),('var_s','.2f'),
                         ('std_s','.2f'),('min','d'),('max','d'),('range','d'),('iqr','d')]:
            print_diff(gs_g, as_g, key, fmt)
        print("\n▼ 正文")
        for key, fmt in [('mean','.2f'),('median','.1f'),('var_s','.2f'),
                         ('std_s','.2f'),('min','d'),('max','d'),('range','d'),('iqr','d')]:
            print_diff(gs_s, as_s, key, fmt)

        # ---- ヒストグラム ----
        print(f"\n{SEP}")
        print(" 文字数ヒストグラム（10文字幅）")
        print(SEP)
        print_hist(gs_g, 'goyo 誤文')
        print_hist(as_g, '人工 誤文')
        print_hist(gs_s, 'goyo 正文')
        print_hist(as_s, '人工 正文')

        # ---- サブタイプ比較 ----
        print(f"\n{SEP}")
        print(" サブタイプ比較（goyo vs 人工データ）")
        print(SEP)
        print_subtype_compare(goyo_recs, art_recs)
    else:
        # goyo にADJがない場合は人工データのみ表示
        print(f"\n{SEP}")
        print(" 文字数ヒストグラム（10文字幅）")
        print(SEP)
        print_hist(as_g, '人工 誤文')
        print_hist(as_s, '人工 正文')

    # ---- サブタイプ分布（人工データ）----
    print(f"\n{SEP}")
    print(" サブタイプ分布（人工データ 100件）")
    print(SEP)
    counts = {}
    nos    = {}
    for r in art_recs:
        st = classify_adj(r['err_tok'], r['cor_tok'])
        counts[st] = counts.get(st, 0) + 1
        nos.setdefault(st, []).append(r['no'])

    total = sum(counts.values())
    for st, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        sample = ','.join(nos[st][:6])
        ellip  = '...' if len(nos[st]) > 6 else ''
        print(f"  {st:<36}: {cnt:3d}件 ({cnt/total*100:4.1f}%)  No.{sample}{ellip}")


if __name__ == '__main__':
    main()
