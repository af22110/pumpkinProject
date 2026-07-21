"""
goyo_extracted（NAIST誤用コーパス）と NEG_v1.tsv の
否定誤り(NEG)データを比較する統計分析スクリプト
"""
import sys, csv, statistics
import openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. データ読み込み
# ============================================================

def load_goyo_neg(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = str(row[3]).lower() if row[3] else ''
        if t.startswith('neg'):
            err_text = row[6]
            cor_text = row[7]
            err_tok  = row[4]
            cor_tok  = str(row[5]) if row[5] else ''
            if err_text is not None and cor_text is not None:
                records.append({
                    'err': str(err_text), 'cor': str(cor_text),
                    'err_tok': err_tok, 'cor_tok': cor_tok
                })
    wb.close()
    return records


def load_artificial_neg(tsv_path):
    records = []
    with open(tsv_path, encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter='\t')
        next(reader, None)
        for row in reader:
            if len(row) < 6:
                continue
            t = str(row[1]).lower() if row[1] else ''
            if t == 'neg':
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

def classify_neg(err_tok, cor_tok, err_sent=''):
    e = str(err_tok) if err_tok is not None else ''
    c = str(cor_tok) if cor_tok is not None else ''
    s = str(err_sent) if err_sent else ''

    # 1. 述語脱落（neg/om）: err_tok が None・空・句断片（。終わり）
    if err_tok is None or e in ('None', ''):
        return '述語脱落（neg/om）'
    if e.endswith('。'):
        return '述語脱落（neg/om）'

    # 2. 二重否定過剰使用
    if any(x in e for x in ['ないこともない', 'なくはない', 'ないわけでは', 'なくもない']):
        return '二重否定過剰使用'

    # 3. ように→ないように（否定脱落）
    if c.endswith('ないように'):
        return 'ように→ないように（否定脱落）'

    # 4. なくて/ないで→ず形
    if any(x in c for x in ['ずに', 'ず、', '問わず', 'せずに', 'ことなしに', 'ずにすみ']) \
       or c.endswith('ず') or c.endswith('ずに'):
        return 'なくて/ないで→ず形'

    # 5. ないで→なくて（付帯状況を原因で書いた誤り）
    if e.endswith('ないで') and (c.endswith('なくて') or c.endswith('なく')):
        return 'ないで→なくて（付帯→原因）'

    # 6. なくて→ないで（原因を付帯状況で書いた誤り）
    if (e.endswith('なくて') or e.endswith('なく')) and c.endswith('ないで'):
        return 'なくて→ないで（原因→付帯）'

    # 7/8. 肯定形→否定形: err_sent の否定呼応副詞の有無で分類
    neg_adverbs = ['しか', 'めったに', 'ほとんど', '全然', 'あまり', '少しも',
                   '二度と', '全く', 'まったく', 'なかなか', 'ほぼ']

    pos_suf = ('ます', 'います', 'します', 'られます', 'えます', 'きます', 'できます',
               'ります', 'みます', 'にます',
               'ました', 'いました', 'しました', 'られました', 'えました', 'きました',
               'できました', 'りました', 'みました')
    neg_suf = ('ません', 'いません', 'しません', 'られません', 'えません', 'きません',
               'できません', 'りません', 'みません',
               'ませんでした', 'いませんでした', 'しませんでした', 'られませんでした',
               'えませんでした', 'きませんでした', 'できませんでした', 'りませんでした',
               'みませんでした')
    neg_end = ('ない', 'ない。', 'なかった', 'ません', 'ません。',
               'ではない', 'ではなかった', 'ではなくて', 'ではありません',
               'ませんでした')

    is_pos_to_neg = (
        (any(e.endswith(x) for x in pos_suf) and any(c.endswith(x) for x in neg_suf))
        or any(c.endswith(x) for x in neg_end)
        or 'ではなかった' in c or 'わない' in c
    )

    if is_pos_to_neg:
        if any(adv in s for adv in neg_adverbs):
            return 'しか/副詞＋肯定形（→否定）'
        return '否定マーカー欠落（肯定→否定）'

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
    def count_subtypes(recs):
        counts = {}
        for r in recs:
            st = classify_neg(r['err_tok'], r['cor_tok'], r.get('err', ''))
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

    print(f"\n  {'サブタイプ':<34}  {'goyo':>18}   {'人工':>14}   差")
    print(f"  {'-'*34}  {'-'*18}   {'-'*14}   {'-'*6}")
    for st in all_types:
        g_n   = g_counts.get(st, 0)
        a_n   = a_counts.get(st, 0)
        g_pct = g_n / g_total * 100
        a_pct = a_n / a_total * 100
        diff  = a_pct - g_pct
        sign  = '+' if diff >= 0 else ''
        print(f"  {st:<34}  {g_n:5d}件({g_pct:5.1f}%)   {a_n:3d}件({a_pct:5.1f}%)   {sign}{diff:.1f}")


# ============================================================
# 5. メイン
# ============================================================

def main():
    base      = Path(__file__).parent.parent
    goyo_path = base / 'goyo_extracted.xlsx'
    tsv_path  = base / 'tsv' / 'NEG_v1.tsv'

    print("データ読み込み中...")
    goyo_recs = load_goyo_neg(goyo_path)
    art_recs  = load_artificial_neg(tsv_path)

    if not goyo_recs:
        print("ERROR: goyo_extracted に NEG型データが見つかりません")
        return
    if not art_recs:
        print("ERROR: NEG_v1.tsv に NEG型データが見つかりません")
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
    print(f" NAIST誤用コーパス（goyo_extracted）: NEG型  {gs_g['n']}件")
    print(SEP)
    print_stats(gs_g, '誤文')
    print_stats(gs_s, '正文')

    # ---- 人工データ ----
    print(f"\n{SEP}")
    print(f" 人工データ（NEG_v1.tsv）: NEG型  {as_g['n']}件")
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

    # ---- サブタイプ（人工データ）----
    print(f"\n{SEP}")
    print(" サブタイプ分布（人工データ 100件）")
    print(SEP)
    counts = {}
    nos    = {}
    for r in art_recs:
        st = classify_neg(r['err_tok'], r['cor_tok'], r.get('err', ''))
        counts[st] = counts.get(st, 0) + 1
        nos.setdefault(st, []).append(r['no'])

    total = sum(counts.values())
    for st, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        sample = ','.join(nos[st][:6])
        ellip  = '...' if len(nos[st]) > 6 else ''
        print(f"  {st:<36}: {cnt:3d}件 ({cnt/total*100:4.1f}%)  No.{sample}{ellip}")

    # ---- サブタイプ比較 ----
    print(f"\n{SEP}")
    print(" サブタイプ比較（goyo vs 人工データ）")
    print(f" ※ goyo の err_tok が単語形のため分類精度は参考値")
    print(SEP)
    print_subtype_compare(goyo_recs, art_recs)


if __name__ == '__main__':
    main()
