import openpyxl
import statistics

wb = openpyxl.load_workbook("goyo_extracted.xlsx", read_only=True)
ws = wb.active

goyo_lengths = []
sei_lengths = []

for row in ws.iter_rows(min_row=2, values_only=True):
    goyo = row[6]  # 誤文
    sei = row[7]   # 正文
    if goyo is not None:
        goyo_lengths.append(len(goyo))
    if sei is not None:
        sei_lengths.append(len(sei))

wb.close()

goyo_var = statistics.variance(goyo_lengths)
sei_var = statistics.variance(sei_lengths)
goyo_pvar = statistics.pvariance(goyo_lengths)
sei_pvar = statistics.pvariance(sei_lengths)

print(f"=== 誤文 (n={len(goyo_lengths)}) ===")
print(f"  最小文字数  : {min(goyo_lengths)}")
print(f"  最大文字数  : {max(goyo_lengths)}")
print(f"  平均文字数  : {statistics.mean(goyo_lengths):.2f}")
print(f"  標本分散    : {goyo_var:.2f}")
print(f"  母分散      : {goyo_pvar:.2f}")
print(f"  標本標準偏差: {statistics.stdev(goyo_lengths):.2f}")

print()
print(f"=== 正文 (n={len(sei_lengths)}) ===")
print(f"  最小文字数  : {min(sei_lengths)}")
print(f"  最大文字数  : {max(sei_lengths)}")
print(f"  平均文字数  : {statistics.mean(sei_lengths):.2f}")
print(f"  標本分散    : {sei_var:.2f}")
print(f"  母分散      : {sei_pvar:.2f}")
print(f"  標本標準偏差: {statistics.stdev(sei_lengths):.2f}")
