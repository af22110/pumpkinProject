"""
goyo_extracted（NAIST誤用コーパス）と v_complete_100.tsv の
動詞誤り(V)データを比較する統計分析スクリプト
"""
import sys, csv, statistics
import openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. データ読み込み
# ============================================================

def load_goyo_v(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = str(row[3]).lower() if row[3] else ''
        if t.startswith('v') and not t.startswith('vt'):
            err_text = row[6]
            cor_text = row[7]
            err_tok  = str(row[4]) if row[4] else ''
            cor_tok  = str(row[5]) if row[5] else ''
            if err_text is not None and cor_text is not None:
                records.append({
                    'err': str(err_text), 'cor': str(cor_text),
                    'err_tok': err_tok, 'cor_tok': cor_tok
                })
    wb.close()
    return records


def load_artificial_v(tsv_path):
    records = []
    with open(tsv_path, encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter='\t')
        next(reader, None)   # ヘッダースキップ
        for row in reader:
            if len(row) < 6:
                continue
            t = str(row[1]).lower() if row[1] else ''
            if t == 'v':
                records.append({
                    'no': row[0],
                    'err_tok': row[2], 'cor_tok': row[3],
                    'err': row[4], 'cor': row[5]
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

IRU_SET = {'いる', 'ある', 'います', 'あります', 'いた', 'あった',
           'いません', 'ありません', 'いて', 'あって'}

TERU = ('ている', 'ていた', 'ています', 'ていません', 'ていて', 'ていました',
        'でいる', 'でいた', 'でいます', 'でいました', 'でいて')

TEKI = ('てきた', 'てきて', 'てきている', 'てきます',
        'ていく', 'ておく', 'ておき', 'てしまっ', 'てしまい')

JUJU = ['くれ', 'あげ', 'もらっ', 'もらい', 'いただ']

PASSIVE = ['られ', 'させ', 'なせ', 'され', 'われ', 'まれ']

HONORIFIC = ['いらっしゃ', 'なさっ', 'おっしゃ']

TE_WRONG = ['いて', 'きて', 'まて', 'みて', 'りて', 'びて']


def has_teru(tok):
    return any(x in tok for x in TERU)

def has_teki(tok):
    return any(x in tok for x in TEKI)

def has_passive(tok):
    return any(x in tok for x in PASSIVE)


def classify_v(err_tok, cor_tok):
    e, c = err_tok, cor_tok

    # 1. いる・ある混同
    if e in IRU_SET and c in IRU_SET and e != c:
        return 'いる・ある混同'

    # 2. ラ抜き言葉（potential without ら）
    if ('られ' in c and 'られ' not in e
            and any(x in e for x in ('れる', 'れた', 'れない', 'れます'))):
        return 'ラ抜き言葉'

    # 3. 授受動詞誤り（くれる・あげる・もらう 混同）
    if any(j in e for j in JUJU) or any(j in c for j in JUJU):
        return '授受動詞誤り'

    # 4. テイル形誤り（〜ている の過不足）
    if has_teru(c) and not has_teru(e):
        return 'テイル形誤り'
    if has_teru(e) and not has_teru(c):
        return 'テイル形誤り'

    # 5. アスペクト誤り（〜てしまう / 〜てくる / 〜ておく 等）
    if has_teki(c) and not has_teki(e):
        return 'アスペクト誤り'
    if has_teki(e) and not has_teki(c):
        return 'アスペクト誤り'

    # 6. テンス誤り（非過去 ↔ 過去）
    if ((e.endswith('ます') and c.endswith('ました'))
            or (e.endswith('ました') and c.endswith('ます'))):
        return 'テンス誤り'
    if ((e.endswith('ない') and c.endswith('なかった'))
            or (e.endswith('なかった') and c.endswith('ない'))):
        return 'テンス誤り'

    # 7. 受身・使役・敬語誤り
    if any(h in c for h in HONORIFIC) or any(h in e for h in HONORIFIC):
        return '受身・使役・敬語誤り'
    if has_passive(c) != has_passive(e):
        return '受身・使役・敬語誤り'

    # 8. テ形誤り（音便テ形の誤活用）
    if any(p in e for p in TE_WRONG) and not any(p in c for p in TE_WRONG):
        return 'テ形誤り'

    return 'その他'


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
    """サブタイプ比率の goyo vs 人工データ比較を表示"""
    def count_subtypes(recs):
        counts = {}
        for r in recs:
            st = classify_v(r['err_tok'], r['cor_tok'])
            counts[st] = counts.get(st, 0) + 1
        return counts

    g_counts = count_subtypes(goyo_recs)
    a_counts = count_subtypes(art_recs)

    # 全サブタイプを列挙（両方の union）
    all_types = sorted(
        set(g_counts) | set(a_counts),
        key=lambda x: -a_counts.get(x, 0)
    )

    g_total = sum(g_counts.values())
    a_total = sum(a_counts.values())

    print(f"\n  {'サブタイプ':<20}  {'goyo':>18}   {'人工':>14}   差")
    print(f"  {'-'*20}  {'-'*18}   {'-'*14}   {'-'*6}")
    for st in all_types:
        g_n   = g_counts.get(st, 0)
        a_n   = a_counts.get(st, 0)
        g_pct = g_n / g_total * 100
        a_pct = a_n / a_total * 100
        diff  = a_pct - g_pct
        sign  = '+' if diff >= 0 else ''
        print(f"  {st:<20}  {g_n:5d}件({g_pct:5.1f}%)   {a_n:3d}件({a_pct:5.1f}%)   {sign}{diff:.1f}")


# ============================================================
# 5. メイン
# ============================================================

def main():
    base      = Path(__file__).parent.parent
    goyo_path = base / 'goyo_extracted.xlsx'
    tsv_path  = base / 'v_complete_100.tsv'

    print("データ読み込み中...")
    goyo_recs = load_goyo_v(goyo_path)
    art_recs  = load_artificial_v(tsv_path)

    if not goyo_recs:
        print("ERROR: goyo_extracted に V型データが見つかりません")
        return
    if not art_recs:
        print("ERROR: v_complete_100.tsv に V型データが見つかりません")
        return

    goyo_err_len = [len(r['err']) for r in goyo_recs]
    goyo_cor_len = [len(r['cor']) for r in goyo_recs]
    art_err_len  = [len(r['err']) for r in art_recs]
    art_cor_len  = [len(r['cor']) for r in art_recs]

    gs_g = calc_stats(goyo_err_len)
    gs_s = calc_stats(goyo_cor_len)
    as_g = calc_stats(art_err_len)
    as_s = calc_stats(art_cor_len)

    # ---- goyo ----
    print(f"\n{SEP}")
    print(f" NAIST誤用コーパス（goyo_extracted）: 動詞誤り(V)型  {gs_g['n']}件")
    print(SEP)
    print_stats(gs_g, '誤文')
    print_stats(gs_s, '正文')

    # ---- 人工データ ----
    print(f"\n{SEP}")
    print(f" 人工データ（v_complete_100.tsv）: 動詞誤り(V)型  {as_g['n']}件")
    print(SEP)
    print_stats(as_g, '誤文')
    print_stats(as_s, '正文')

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

    # ---- サブタイプ ----
    print(f"\n{SEP}")
    print(" サブタイプ分布（人工データ 100件）")
    print(SEP)
    counts = {}
    nos    = {}
    for r in art_recs:
        st = classify_v(r['err_tok'], r['cor_tok'])
        counts[st] = counts.get(st, 0) + 1
        nos.setdefault(st, []).append(r['no'])

    total = sum(counts.values())
    for st, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        sample = ','.join(nos[st][:6])
        ellip  = '...' if len(nos[st]) > 6 else ''
        print(f"  {st:<20}: {cnt:3d}件 ({cnt/total*100:4.1f}%)  No.{sample}{ellip}")

    # ---- サブタイプ比較 ----
    print(f"\n{SEP}")
    print(" サブタイプ比較（goyo vs 人工データ）")
    print(f" ※ goyo の err_tok/cor_tok が空欄の場合は「その他」扱い")
    print(SEP)
    print_subtype_compare(goyo_recs, art_recs)


if __name__ == '__main__':
    main()
