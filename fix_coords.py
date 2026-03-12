import csv

K1 =  1269.3
K2 = -867.6

path = r'C:\Projects\arborist-plans\Projects\7631-creditview\data.csv'
rows = []
with open(path, newline='', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    fields = reader.fieldnames
    for row in reader:
        n = int(row['tree_num']) if row['tree_num'].isdigit() else 0
        if n >= 43 and row['cx'] and row['cy']:
            cx, cy = float(row['cx']), float(row['cy'])
            new_cx = round(K1 - cx, 2)
            new_cy = round(K2 - cy, 2)
            print(f"Tree {n}: ({cx}, {cy}) -> ({new_cx}, {new_cy})")
            row['cx'] = str(new_cx)
            row['cy'] = str(new_cy)
        rows.append(row)

with open(path, 'w', newline='', encoding='utf-8') as f:
    w = csv.DictWriter(f, fieldnames=fields)
    w.writeheader()
    w.writerows(rows)

print("Done.")
