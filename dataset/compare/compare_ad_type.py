"""
goyo_extracted（NAIST誤用コーパス）と ad_batch1_50.tsv の
余剰誤り(AD)データを比較する統計分析スクリプト
"""
import sys, csv, statistics
import openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. データ読み込み
# ============================================================

def load_goyo_ad(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = str(row[3]).lower() if row[3] else ''
        if t.startswith('ad'):
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


def load_artificial_ad(tsv_path):
    records = []
    with open(tsv_path, encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter='\t')
        for row in reader:
            if len(row) < 6:
                continue
            t = str(row[1]).lower() if row[1] else ''
            if t == 'ad':
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

# お/ご + カタカナ・外来語（本来は丁寧語が不要な語）
def _is_okeigo(err_tok, cor_tok):
    if not err_tok:
        return False
    if err_tok.startswith('お') and not cor_tok.startswith('お'):
        return True
    if err_tok.startswith('ご') and not cor_tok.startswith('ご'):
        return True
    return False

# 指示詞余剰
SHIJI = ['この', 'その', 'あの', 'こんな', 'そんな', 'あんな',
         'これ', 'それ', 'あれ', 'こちら', 'そちら', 'あちら']

def _is_shiji(err_tok, cor_tok):
    for s in SHIJI:
        if err_tok.startswith(s) and not cor_tok.startswith(s):
            return True
    return False

# 余剰副詞・接続表現
FUKUSHI = ['でも', 'もし', 'やっぱり', 'もちろん', '一生懸命', 'だんだん',
           'ちゃんと', 'ぜひ', 'いつも', 'ぜったい', 'まず', '正直に言うと',
           'たまたま', '実は', '先週', 'この間', 'きっと', 'きちんと']

def _is_fukushi(err_tok, cor_tok):
    for f in FUKUSHI:
        if f in err_tok and f not in cor_tok:
            return True
    return False

# 余剰動作節 (〜てから)
def _is_dousa_setsu(err_tok, cor_tok):
    return 'てから' in err_tok and 'てから' not in cor_tok

# の/こと余剰
def _is_nokoto(err_tok, cor_tok):
    for pat in ['のこと', 'ことの', 'ことを', 'ことが']:
        if pat in err_tok and pat not in cor_tok:
            return True
    return False

# のだ余剰
def _is_noda(err_tok, cor_tok):
    for pat in ['のだ', 'のです', 'なのだ', 'なのです', 'んです', 'んだ']:
        if pat in err_tok and pat not in cor_tok:
            return True
    return False


def classify_ad(err_tok, cor_tok):
    e, c = err_tok, cor_tok
    if _is_okeigo(e, c):
        return 'お/ご過剰付与'
    if _is_shiji(e, c):
        return '指示詞余剰'
    if _is_dousa_setsu(e, c):
        return '余剰動作節'
    if _is_nokoto(e, c):
        return 'の/こと余剰'
    if _is_noda(e, c):
        return 'のだ余剰'
    if _is_fukushi(e, c):
        return '余剰副詞・接続'
    # err_tokが空でもcor_tokが空: 削除パターン → 修飾語余剰（推定）
    if e and not c:
        return '修飾語余剰（全削除）'
    return 'その他（修飾語余剰等）'


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
            st = classify_ad(r['err_tok'], r['cor_tok'])
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

    print(f"\n  {'サブタイプ':<22}  {'goyo':>20}   {'人工':>14}   差")
    print(f"  {'-'*22}  {'-'*20}   {'-'*14}   {'-'*6}")
    for st in all_types:
        g_n   = g_counts.get(st, 0)
        a_n   = a_counts.get(st, 0)
        g_pct = g_n / g_total * 100
        a_pct = a_n / a_total * 100
        diff  = a_pct - g_pct
        sign  = '+' if diff >= 0 else ''
        print(f"  {st:<22}  {g_n:5d}件({g_pct:5.1f}%)   {a_n:3d}件({a_pct:5.1f}%)   {sign}{diff:.1f}")


# ============================================================
# 5. メイン
# ============================================================

def main():
    base      = Path(__file__).parent.parent
    goyo_path = base / 'goyo_extracted.xlsx'
    tsv_path  = base / 'tsv' / 'ad_batch1_50.tsv'

    print("データ読み込み中...")
    goyo_recs = load_goyo_ad(goyo_path)
    art_recs  = load_artificial_ad(tsv_path)

    if not goyo_recs:
        print("ERROR: goyo_extracted に AD型データが見つかりません")
        return
    if not art_recs:
        print("ERROR: ad_batch1_50.tsv に AD型データが見つかりません")
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
    print(f" NAIST誤用コーパス（goyo_extracted）: 余剰誤り(AD)型  {gs_g['n']}件")
    print(SEP)
    print_stats(gs_g, '誤文')
    print_stats(gs_s, '正文')

    # ---- 人工データ ----
    print(f"\n{SEP}")
    print(f" 人工データ（ad_batch1_50.tsv）: 余剰誤り(AD)型  {as_g['n']}件")
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
        st = classify_ad(r['err_tok'], r['cor_tok'])
        counts[st] = counts.get(st, 0) + 1
        nos.setdefault(st, []).append(r['no'])

    total = sum(counts.values())
    for st, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        sample = ','.join(nos[st][:6])
        ellip  = '...' if len(nos[st]) > 6 else ''
        print(f"  {st:<24}: {cnt:3d}件 ({cnt/total*100:4.1f}%)  No.{sample}{ellip}")

    # ---- サブタイプ比較 ----
    print(f"\n{SEP}")
    print(" サブタイプ比較（goyo vs 人工データ）")
    print(f" ※ goyo の err_tok/cor_tok が空欄の場合は「その他」扱い")
    print(SEP)
    print_subtype_compare(goyo_recs, art_recs)


if __name__ == '__main__':
    main()
