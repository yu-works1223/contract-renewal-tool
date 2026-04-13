import pandas as pd

# CSVを読み込む
df = pd.read_csv("tenants.csv")

# datetime型に変換
df["contract_start"] = pd.to_datetime(df["contract_start"])

# 2年後の更新日を計算
df["renewal_date"] = df["contract_start"] + pd.DateOffset(years=2)

# 更新年・更新月を列として追加
df["renewal_year"] = df["renewal_date"].dt.year
df["renewal_month"] = df["renewal_date"].dt.month

# 年・月ごとに件数を集計
summary = df.groupby(["renewal_year", "renewal_month"]).size().reset_index(name="count")

print("=== 更新月別 件数集計 ===")
print(summary)
