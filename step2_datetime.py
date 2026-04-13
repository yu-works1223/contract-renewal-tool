import pandas as pd

# CSVを読み込む
df = pd.read_csv("tenants.csv")

# str型 → datetime型に変換
df["contract_start"] = pd.to_datetime(df["contract_start"])

# 2年後の更新日を計算
df["renewal_date"] = df["contract_start"] + pd.DateOffset(years=2)

# 確認
print("=== 更新日計算結果 ===")
print(df[["tenant_name", "contract_start", "renewal_date"]])

print("\n=== データ型 ===")
print(df[["contract_start", "renewal_date"]].dtypes)
