import pandas as pd

# CSVを読み込む
df = pd.read_csv("tenants.csv")

# datetime型に変換
df["contract_start"] = pd.to_datetime(df["contract_start"])

# 2年後の更新日を計算
df["renewal_date"] = df["contract_start"] + pd.DateOffset(years=2)

# 更新年・更新月を追加
df["renewal_year"] = df["renewal_date"].dt.year
df["renewal_month"] = df["renewal_date"].dt.month

# 月別集計
summary = df.groupby(["renewal_year", "renewal_month"]).size().reset_index(name="count")

# Excelに書き出す（2シート）
with pd.ExcelWriter("契約更新リスト.xlsx", engine="openpyxl") as writer:
    df[["tenant_id", "tenant_name", "room_number", "contract_start", "renewal_date"]].to_excel(
        writer, sheet_name="一覧", index=False
    )
    summary.to_excel(writer, sheet_name="集計", index=False)

print("=== 完了 ===")
print("契約更新リスト.xlsx を出力しました！")
