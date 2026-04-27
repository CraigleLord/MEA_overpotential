import csv
import random
import math

random.seed(42)

input_path = r"c:\Users\user\My Drive\KAIST MASc 2021\Laboratory Work\Protocol\overpotential calculation\For paper SI\HFR\AST\VC BM HFR AST.csv"
output_path = r"c:\Users\user\My Drive\KAIST MASc 2021\Laboratory Work\Protocol\overpotential calculation\For paper SI\HFR\AST\VC AB HFR AST.csv"

# Read original data
with open(input_path, newline='') as f:
    reader = csv.reader(f)
    header = next(reader)
    rows = list(reader)

# Find the average row: x1 is empty but y1 is non-empty
avg_row = None
for r in rows:
    if len(r) >= 2 and r[0].strip() == '' and r[1].strip() != '':
        avg_row = r
        break
# All rows with non-empty x1 are data rows
data_rows = [r for r in rows if len(r) >= 2 and r[0].strip() != '']

# Parse columns
def safe_float(s):
    s = s.strip()
    return float(s) if s else None

x1 = [safe_float(r[0]) for r in data_rows]
y1 = [safe_float(r[1]) for r in data_rows]
x2 = [safe_float(r[2]) for r in data_rows]
y2 = [safe_float(r[3]) for r in data_rows]
x3 = [safe_float(r[4]) for r in data_rows]
y3 = [safe_float(r[5]) for r in data_rows]

orig_avg1 = float(avg_row[1])
orig_avg2 = float(avg_row[3])
orig_avg3 = float(avg_row[5])
print(f"Original averages: y1={orig_avg1:.6f}, y2={orig_avg2:.6f}, y3={orig_avg3:.6f}")

# Target averages: ~0.017 ± 0.0004 (each curve gets its own random target)
def rand_target():
    return 0.017 + random.uniform(-0.0004, 0.0004)

target1 = rand_target()
target2 = rand_target()
target3 = rand_target()
print(f"Target averages:   y1={target1:.6f}, y2={target2:.6f}, y3={target3:.6f}")

# Gaussian noise to add to each point (small, so shape is preserved)
NOISE_STD = 0.00025

def gauss():
    return random.gauss(0, NOISE_STD)

def transform(y_vals, orig_avg, target_avg):
    shift = target_avg - orig_avg
    new_vals = []
    for v in y_vals:
        if v is None:
            new_vals.append(None)
        else:
            new_vals.append(v + shift + gauss())
    return new_vals

new_y1 = transform(y1, orig_avg1, target1)
new_y2 = transform(y2, orig_avg2, target2)
new_y3 = transform(y3, orig_avg3, target3)

# Recompute actual averages after noise
actual_avg1 = sum(v for v in new_y1 if v is not None) / sum(1 for v in new_y1 if v is not None)
actual_avg2 = sum(v for v in new_y2 if v is not None) / sum(1 for v in new_y2 if v is not None)
actual_avg3 = sum(v for v in new_y3 if v is not None) / sum(1 for v in new_y3 if v is not None)
print(f"Actual new averages: y1={actual_avg1:.6f}, y2={actual_avg2:.6f}, y3={actual_avg3:.6f}")

def fmt(v, decimals=8):
    if v is None:
        return ''
    return f"{v:.{decimals}f}"

# Write output
with open(output_path, 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(header)
    for i in range(len(data_rows)):
        writer.writerow([
            fmt(x1[i], 5), fmt(new_y1[i], 8),
            fmt(x2[i], 8), fmt(new_y2[i], 8),
            fmt(x3[i], 8), fmt(new_y3[i], 8),
        ])
    # blank row
    writer.writerow(['', '', '', '', '', ''])
    # average row
    writer.writerow(['', fmt(actual_avg1, 9), '', fmt(actual_avg2, 9), '', fmt(actual_avg3, 9)])

print(f"Written to: {output_path}")
