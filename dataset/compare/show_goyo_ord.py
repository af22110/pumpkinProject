"""goyo_extracted の ORD型エントリーをサンプル表示"""
import sys, openpyxl
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

path = Path(__file__).parent.parent / 'goyo_extracted.xlsx'
wb = openpyxl.load_workbook(path, read_only=True)
ws = wb.active

records = []
for row in ws.iter_rows(min_row=2, values_only=True):
    t = str(row[3]).lower() if row[3] else ''
    if t.startswith('ord'):
        records.append({
            'type':    str(row[3]),
            'err_tok': str(row[4]) if row[4] else '',
            'cor_tok': str(row[5]) if row[5] else '',
            'err':     str(row[6]) if row[6] else '',
            'cor':     str(row[7]) if row[7] else '',
        })
wb.close()

print(f"ORD型 合計: {len(records)}件\n")
print(f"{'#':<4} {'err_tok':<20} {'cor_tok':<20} {'err_sent'}")
print('-'*90)
for i, r in enumerate(records, 1):
    print(f"{i:<4} {r['err_tok']:<20} {r['cor_tok']:<20} {r['err']}")
