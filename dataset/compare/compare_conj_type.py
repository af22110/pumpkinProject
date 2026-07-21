"""
goyo_extracted（NAIST誤用コーパス）と conj_complete_100.tsv の
接続表現誤り(CONJ)データを比較する統計分析スクリプト
"""
import sys, csv, statistics
import openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. データ読み込み
# ============================================================

def load_goyo_conj(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = str(row[3]).lower() if row[3] else ''
        if t in ('conj', 'conj/part'):
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


def load_artificial_conj(tsv_path):
    records = []
    with open(tsv_path, encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter='\t')
        for row in reader:
            if len(row) < 6:
                continue
            t = str(row[1]).lower() if row[1] else ''
            if t == 'conj':
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

CONTRAST  = {'しかし','しかしながら','が','でも','ところが','なのに','のに',
             'にもかかわらず','ものの','だが','それなのに','けれど','けれども',
             'けど','だのに','だけれど','だけれども','いっぽう'}
CAUSAL    = {'だから','ので','から','そのため','したがって','ゆえに','よって',
             'それで','なので','ですから','それゆえ','このため','それゆえに',
             'のことで','のためだ','のため','なので'}
ADDITIVE  = {'それに','また','そのうえ','そして','それから','さらに'}
COND_FORMS = {'たら','ても','と','ながら','ている間','いながら','ば','なら',
              'したら','たり','ると','れば'}
DISCOURSE = {'さて','ところで','ところが','このように','それから'}


def classify_conj(err_tok, cor_tok):
    e, c = err_tok.strip(), cor_tok.strip()

    # 条件・様態形式の混同（たら/ても/と/ながら/間 etc.）
    cond_e = any(f in e for f in COND_FORMS)
    cond_c = any(f in c for f in COND_FORMS)
    if cond_e or cond_c:
        return '条件・様態形式の混同'

    # 逆接と順接の混同
    e_cont = any(f in e for f in CONTRAST)
    e_caus = any(f in e for f in CAUSAL)
    c_cont = any(f in c for f in CONTRAST)
    c_caus = any(f in c for f in CAUSAL)
    if (e_cont and c_caus) or (e_caus and c_cont):
        return '逆接と順接の混同'

    # 累加・列挙表現の混同（そして/それに/また/そのうえ）
    add_e = any(f in e for f in ADDITIVE)
    add_c = any(f in c for f in ADDITIVE)
    if add_e and add_c:
        return '累加・列挙表現の混同'

    # 談話標識の混同（さて/ところで/ところが）
    disc_e = any(f in e for f in DISCOURSE)
    disc_c = any(f in c for f in DISCOURSE)
    if disc_e or disc_c:
        return '談話標識の混同'

    # 形式・文体差（けど↔けれども 等）
    informal = {'けど', 'だけど', 'いいけど', 'いるけど'}
    formal   = {'けれど', 'けれども', 'だけれど', 'だけれども', 'ものの',
                'にもかかわらず', 'しかしながら'}
    if (any(f in e for f in informal) and any(f in c for f in formal)) or \
       (any(f in e for f in formal)   and any(f in c for f in informal)):
        return '形式・文体差'

    # 同一方向内の順接混同（それで↔したがって 等）
    if e_caus and c_caus:
        return '順接表現間の混同'

    # 同一方向内の逆接混同
    if e_cont and c_cont:
        return '逆接表現間の混同'

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
            st = classify_conj(r['err_tok'], r['cor_tok'])
            counts[st] = counts.get(st, 0) + 1
        return counts

    g_counts = count_subtypes(goyo_recs)
    a_counts = count_subtypes(art_recs)
    all_types = sorted(set(g_counts) | set(a_counts), key=lambda x: -a_counts.get(x, 0))
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
    tsv_path  = base / 'tsv' / 'conj_complete_100.tsv'

    if not tsv_path.exists():
        tsv_path = base / 'conj_complete_100.tsv'

    print("データ読み込み中...")
    goyo_recs = load_goyo_conj(goyo_path)
    art_recs  = load_artificial_conj(tsv_path)

    if not goyo_recs:
        print("ERROR: goyo_extracted に CONJ型データが見つかりません")
        return
    if not art_recs:
        print("ERROR: conj_complete_100.tsv に CONJ型データが見つかりません")
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
    print(f" NAIST誤用コーパス（goyo_extracted）: CONJ型(conj+conj/part)  {gs_g['n']}件")
    print(SEP)
    print_stats(gs_g, '誤文')
    print_stats(gs_s, '正文')

    print(f"\n{SEP}")
    print(f" 人工データ（conj_complete_100.tsv）: CONJ型  {as_g['n']}件")
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
    print(" サブタイプ分布（人工データ 100件）")
    print(SEP)
    counts, nos = {}, {}
    for r in art_recs:
        st = classify_conj(r['err_tok'], r['cor_tok'])
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
