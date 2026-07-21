"""
goyo_extracted（NAIST誤用コーパス）と nom_batch1_30.tsv の
名詞化誤り(NOM)データを比較する統計分析スクリプト
"""
import sys, csv, statistics
import openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. データ読み込み
# ============================================================

def _is_nokoto(err_tok, cor_tok):
    """の/こと混同パターンかどうか判定"""
    e, c = err_tok.strip(), cor_tok.strip()
    if ('の' in e and 'こと' in c) or ('こと' in e and 'の' in c):
        return True
    if e in ('の', 'のが', 'のは', 'のを', 'のに') and not c:
        return True
    if not e and c in ('こと', 'ことが', 'ことは', 'ことを', 'ことに'):
        return True
    return False


def load_goyo_nom(path):
    """noun/comp 型のうち の/こと混同に絞って読み込む"""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = str(row[3]).lower() if row[3] else ''
        if t == 'noun/comp':
            err_text = row[6]
            cor_text = row[7]
            err_tok  = str(row[4]) if row[4] else ''
            cor_tok  = str(row[5]) if row[5] else ''
            if err_text is not None and cor_text is not None:
                if _is_nokoto(err_tok, cor_tok):
                    records.append({
                        'err': str(err_text), 'cor': str(cor_text),
                        'err_tok': err_tok, 'cor_tok': cor_tok,
                        'type': t,
                    })
    wb.close()
    return records


def load_artificial_nom(tsv_path):
    records = []
    with open(tsv_path, encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter='\t')
        for row in reader:
            if len(row) < 6:
                continue
            t = str(row[1]).lower() if row[1] else ''
            if t == 'nom':
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

def classify_nom(err_tok, cor_tok):
    e, c = err_tok.strip(), cor_tok.strip()

    # のだ / ことだ
    if ('のだ' in e and 'ことだ' in c) or ('ことだ' in e and 'のだ' in c):
        return 'のだ/ことだ混同'

    # のに / ことに（驚き・気づき）
    if ('のに' in e and 'ことに' in c) or ('ことに' in e and 'のに' in c):
        return 'のに/ことに混同（驚き）'

    # の → こと（一般名詞節）
    no_forms   = ('のが', 'のは', 'のを', 'のと', 'のも', 'のだ')
    koto_forms = ('ことが', 'ことは', 'ことを', 'ことと', 'ことも', 'ことだ')
    if e in no_forms and c in koto_forms:
        return 'の→こと（名詞節）'

    # こと → の（知覚動詞・目撃動詞）
    if e in koto_forms and c in no_forms:
        return 'こと→の（知覚動詞）'

    # 単独の / こと
    if e == 'の' and c == 'こと':
        return 'の→こと（名詞節）'
    if e == 'こと' and c == 'の':
        return 'こと→の（知覚動詞）'

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
            st = classify_nom(r['err_tok'], r['cor_tok'])
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

    print(f"\n  {'サブタイプ':<26}  {'goyo':>20}   {'人工':>14}   差")
    print(f"  {'-'*26}  {'-'*20}   {'-'*14}   {'-'*6}")
    for st in all_types:
        g_n   = g_counts.get(st, 0)
        a_n   = a_counts.get(st, 0)
        g_pct = g_n / g_total * 100
        a_pct = a_n / a_total * 100
        diff  = a_pct - g_pct
        sign  = '+' if diff >= 0 else ''
        print(f"  {st:<26}  {g_n:5d}件({g_pct:5.1f}%)   {a_n:3d}件({a_pct:5.1f}%)   {sign}{diff:.1f}")


# ============================================================
# 5. メイン
# ============================================================

def main():
    base      = Path(__file__).parent.parent
    goyo_path = base / 'goyo_extracted.xlsx'

    # ファイル名にスペースが入っている場合も探す
    tsv_candidates = [
        base / 'tsv' / 'nom_batch1_30.tsv',
        base / 'tsv' / 'nom_batch1_30 .tsv',
        base / 'nom_batch1_30.tsv',
        base / 'nom_batch1_30 .tsv',
    ]
    tsv_path = None
    for p in tsv_candidates:
        if p.exists():
            tsv_path = p
            break

    if tsv_path is None:
        print("ERROR: nom_batch1_30.tsv が見つかりません")
        return

    print(f"TSVファイル: {tsv_path.name}")
    print("データ読み込み中...")
    goyo_recs = load_goyo_nom(goyo_path)
    art_recs  = load_artificial_nom(tsv_path)

    if not goyo_recs:
        print("ERROR: goyo_extracted に NOM型(noun/comp の/こと混同)データが見つかりません")
        return
    if not art_recs:
        print("ERROR: nom_batch1_30.tsv に NOM型データが見つかりません")
        return

    goyo_err_len = [len(r['err']) for r in goyo_recs]
    goyo_cor_len = [len(r['cor']) for r in goyo_recs]
    art_err_len  = [len(r['err']) for r in art_recs]
    art_cor_len  = [len(r['cor']) for r in art_recs]

    gs_g = calc_stats(goyo_err_len)
    gs_s = calc_stats(goyo_cor_len)
    as_g = calc_stats(art_err_len)
    as_s = calc_stats(art_cor_len)

    print(f"\n{SEP}")
    print(f" NAIST誤用コーパス（goyo_extracted）: noun/comp(の/こと混同)  {gs_g['n']}件")
    print(SEP)
    print_stats(gs_g, '誤文')
    print_stats(gs_s, '正文')

    print(f"\n{SEP}")
    print(f" 人工データ（nom_batch1_30.tsv）: 名詞化誤り(NOM)型  {as_g['n']}件")
    print(SEP)
    print_stats(as_g, '誤文')
    print_stats(as_s, '正文')

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

    print(f"\n{SEP}")
    print(" 文字数ヒストグラム（10文字幅）")
    print(SEP)
    print_hist(gs_g, 'goyo 誤文')
    print_hist(as_g, '人工 誤文')
    print_hist(gs_s, 'goyo 正文')
    print_hist(as_s, '人工 正文')

    print(f"\n{SEP}")
    print(" サブタイプ分布（人工データ 30件）")
    print(SEP)
    counts = {}
    nos    = {}
    for r in art_recs:
        st = classify_nom(r['err_tok'], r['cor_tok'])
        counts[st] = counts.get(st, 0) + 1
        nos.setdefault(st, []).append(r['no'])

    total = sum(counts.values())
    for st, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        sample = ','.join(nos[st][:6])
        ellip  = '...' if len(nos[st]) > 6 else ''
        print(f"  {st:<26}: {cnt:3d}件 ({cnt/total*100:4.1f}%)  No.{sample}{ellip}")

    print(f"\n{SEP}")
    print(" サブタイプ比較（goyo vs 人工データ）")
    print(f" ※ goyo の err_tok/cor_tok が空欄の場合は「その他」扱い")
    print(SEP)
    print_subtype_compare(goyo_recs, art_recs)


if __name__ == '__main__':
    main()
