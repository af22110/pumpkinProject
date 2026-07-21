"""
goyo_extracted（NAIST誤用コーパス）と PRON_v1_full100.tsv の
代名詞誤り(PRON)データを比較する統計分析スクリプト
"""
import sys, csv, statistics
import openpyxl
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 1. データ読み込み
# ============================================================

def load_goyo_pron(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        t = str(row[3]).lower().strip() if row[3] else ''
        if t == 'pron':
            err_text = row[6]
            cor_text = row[7]
            err_tok  = str(row[4]) if row[4] else ''
            cor_tok  = str(row[5]) if row[5] else ''
            if err_text is not None and cor_text is not None:
                records.append({
                    'err': str(err_text), 'cor': str(cor_text),
                    'err_tok': err_tok, 'cor_tok': cor_tok,
                })
    wb.close()
    return records


def load_artificial_pron(tsv_path):
    records = []
    with open(tsv_path, encoding='utf-8-sig') as f:
        reader = csv.reader(f, delimiter='\t')
        for row in reader:
            if len(row) < 6:
                continue
            # ヘッダー行をスキップ
            if row[0] == 'No':
                continue
            t = str(row[1]).lower().strip() if row[1] else ''
            if t == 'pron':
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

# 数量詞・不定語
QUANTIFIERS = {
    'みんな', '誰にでも', 'どれも', 'どちらも', 'みな',
    '何も', '何でも', '誰でも', '誰も', 'どこでも', 'どこにも',
    'こんなに', 'そんなに', 'そう',
}

# コ/ソ/ア指示詞
KOSOADO = {
    'これ', 'それ', 'あれ', 'この', 'その', 'あの',
    'ここ', 'そこ', 'あそこ', 'こちら', 'そちら', 'あちら',
    'こう', 'そう', 'ああ', 'こんな', 'そんな', 'あんな',
    'これら', 'それら', 'あれら',
    '今日', '明日', '昨日',  # 時制コ/ソ混同を含む
}

# 複数形の過剰適用
PLURAL_MARKERS = ('たち', 'ら', '達')

def is_plural_error(err_tok):
    """複数形過剰（彼たち・ら達・たちら等）の判定"""
    e = err_tok
    if 'ら達' in e or '達ら' in e:
        return True
    if 'たちら' in e or 'らたち' in e:
        return True
    # 「彼たち」「これたち」「それたち」等（たち単独は除外）
    prefixes = ['彼', '彼女', 'これ', 'それ', 'あれ', 'あの人']
    for p in prefixes:
        if e.startswith(p) and 'たち' in e and 'ら' not in e:
            # 「彼女たち」は正しい日本語なので除外、「彼たち」のみ対象
            if p in ('彼', 'これ', 'それ', 'あれ') and 'たち' in e:
                return True
    return False

def classify_pron(err_tok, cor_tok):
    e = str(err_tok).strip() if err_tok else ''
    c = str(cor_tok).strip() if cor_tok else ''

    # 1. 複数形過剰適用（彼たち・ら達・あの人たちら 等）
    if is_plural_error(e):
        return '複数形過剰適用'

    # 2. 数量詞・不定語の選択誤り
    if e in QUANTIFIERS or c in QUANTIFIERS:
        return '数量詞の選択誤り'

    # 3. コ/ソ/ア距離混同（どちらも指示詞）
    if e in KOSOADO and c in KOSOADO:
        return 'コ/ソ/ア距離混同'

    # 4. 代名詞の不使用（NP反復）：err_tokが長い（NP）、cor_tokが指示詞
    if len(e) > 6 and (c in KOSOADO or c in ('彼は', '彼女は', 'その子は')):
        return '代名詞の不使用（NP反復）'
    if len(e) > 6 and len(c) <= 6:
        return '代名詞の不使用（NP反復）'

    # 5. 代名詞の過剰使用（指示詞・人称代名詞を使って指示対象が不明瞭）
    return '代名詞の過剰使用'


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


def print_diff(gs, as_, key, fmt='.2f'):
    gv = gs[key]
    av = as_[key]
    diff = av - gv
    sign = '+' if diff >= 0 else ''
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


# ============================================================
# 5. メイン
# ============================================================

def main():
    base      = Path(__file__).parent.parent
    goyo_path = base / 'goyo_extracted.xlsx'
    tsv_path  = base / 'tsv' / 'PRON_v1_full100.tsv'

    print("データ読み込み中...")
    goyo_recs = load_goyo_pron(goyo_path)
    art_recs  = load_artificial_pron(tsv_path)

    if not goyo_recs:
        print("WARNING: goyo_extracted に PRON型（type='pron'）のデータが見つかりません")
        print("  → goyo データなしで人工データの統計のみ表示します")
    else:
        print(f"goyo PRON: {len(goyo_recs)}件")
    print(f"人工 PRON: {len(art_recs)}件")

    art_err_len = [len(r['err']) for r in art_recs]
    art_cor_len = [len(r['cor']) for r in art_recs]
    as_g = calc_stats(art_err_len)
    as_s = calc_stats(art_cor_len)

    # ---- 人工データ単体 ----
    print(f"\n{SEP}")
    print(f" 人工データ（PRON_v1_full100.tsv）: PRON型  {as_g['n']}件")
    print(SEP)
    print_stats(as_g, '誤文')
    print_stats(as_s, '正文')

    # ---- goyo との比較（goyo があれば）----
    if goyo_recs:
        goyo_err_len = [len(r['err']) for r in goyo_recs]
        goyo_cor_len = [len(r['cor']) for r in goyo_recs]
        gs_g = calc_stats(goyo_err_len)
        gs_s = calc_stats(goyo_cor_len)

        print(f"\n{SEP}")
        print(f" NAIST誤用コーパス（goyo_extracted）: PRON型  {gs_g['n']}件")
        print(SEP)
        print_stats(gs_g, '誤文')
        print_stats(gs_s, '正文')

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

        # ---- ヒストグラム比較 ----
        print(f"\n{SEP}")
        print(" 文字数ヒストグラム（10文字幅）")
        print(SEP)
        print_hist(gs_g, 'goyo 誤文')
        print_hist(as_g, '人工 誤文')
        print_hist(gs_s, 'goyo 正文')
        print_hist(as_s, '人工 正文')
    else:
        # goyo なしでも人工データのヒストグラムは表示
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
        st = classify_pron(r['err_tok'], r['cor_tok'])
        counts[st] = counts.get(st, 0) + 1
        nos.setdefault(st, []).append(r['no'])

    total = sum(counts.values())
    for st, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        sample = ','.join(nos[st][:6])
        ellip  = '...' if len(nos[st]) > 6 else ''
        print(f"  {st:<28}: {cnt:3d}件 ({cnt/total*100:4.1f}%)  No.{sample}{ellip}")

    # ---- サブタイプ比較（goyo があれば）----
    if goyo_recs:
        print(f"\n{SEP}")
        print(" サブタイプ比較（goyo vs 人工データ）")
        print(SEP)

        g_counts = {}
        for r in goyo_recs:
            st = classify_pron(r['err_tok'], r['cor_tok'])
            g_counts[st] = g_counts.get(st, 0) + 1

        all_types = sorted(set(g_counts) | set(counts), key=lambda x: -counts.get(x, 0))
        g_total = sum(g_counts.values())

        print(f"\n  {'サブタイプ':<28}  {'goyo':>16}   {'人工':>14}   差")
        print(f"  {'-'*28}  {'-'*16}   {'-'*14}   {'-'*6}")
        for st in all_types:
            g_n   = g_counts.get(st, 0)
            a_n   = counts.get(st, 0)
            g_pct = g_n / g_total * 100 if g_total else 0
            a_pct = a_n / total   * 100 if total   else 0
            diff  = a_pct - g_pct
            sign  = '+' if diff >= 0 else ''
            print(f"  {st:<28}  {g_n:5d}件({g_pct:5.1f}%)   {a_n:3d}件({a_pct:5.1f}%)   {sign}{diff:.1f}")

    # ---- 整合性チェック ----
    print(f"\n{SEP}")
    print(" 整合性チェック（err_tok出現・cor_sent置換）")
    print(SEP)
    ng = []
    for r in art_recs:
        et = r['err_tok']
        ct = r['cor_tok']
        es = r['err']
        cs = r['cor']
        cnt = es.count(et)
        expected_cs = es.replace(et, ct, 1)
        issues = []
        if cnt != 1:
            issues.append(f'err_tok出現{cnt}回')
        if cs != expected_cs:
            issues.append('cor_sent不一致')
        if issues:
            ng.append((r['no'], et, ct, ', '.join(issues)))

    if ng:
        print(f"  NG: {len(ng)}件")
        for no, et, ct, msg in ng:
            print(f"    No.{no}  {et}→{ct}  {msg}")
    else:
        print(f"  全{len(art_recs)}件 OK（err_tok出現1回・cor_sent置換一致）")


if __name__ == '__main__':
    main()
