import pandas as pd

# CSVを読み込む
df = pd.read_csv("tenants.csv")

# 内容を確認
print("=== カラム名 ===")
print(df.columns.tolist())

print("\n=== 先頭5行 ===")
print(df.head())

print("\n=== データ型 ===")
print(df.dtypes)
