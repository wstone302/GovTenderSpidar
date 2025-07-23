import requests
import pandas as pd
from datetime import datetime
import os

# === 參數設定 ===
api_key = "CWA-0DD9CB08-1E98-46D3-89A0-F4DB88791447"
url = f"https://opendata.cwa.gov.tw/api/v1/rest/datastore/F-C0032-001?Authorization={api_key}"
output_file = "weather_history.xlsx"

# === 抓取資料 ===
response = requests.get(url)
data = response.json()

# === 確認有資料回傳 ===
if "records" not in data:
    print("錯誤：無法取得資料")
    exit()

# === 整理成 DataFrame ===
rows = []
now = datetime.now().strftime("%Y-%m-%d")

for location in data["records"]["location"]:
    city = location["locationName"]
    elements = {e["elementName"]: e["time"] for e in location["weatherElement"]}
    
    for i in range(len(elements["Wx"])):
        rows.append({
            "紀錄日期": now,
            "縣市": city,
            "預報起": elements["Wx"][i]["startTime"],
            "預報迄": elements["Wx"][i]["endTime"],
            "天氣現象": elements["Wx"][i]["parameter"]["parameterName"],
            "降雨機率": elements["PoP"][i]["parameter"]["parameterName"] + "%",
            "最高溫": elements["MaxT"][i]["parameter"]["parameterName"] + "°C",
            "最低溫": elements["MinT"][i]["parameter"]["parameterName"] + "°C",
            "舒適度": elements["CI"][i]["parameter"]["parameterName"]
        })

df_new = pd.DataFrame(rows)

# === 合併並儲存 ===
if os.path.exists(output_file):
    df_old = pd.read_excel(output_file)
    df_combined = pd.concat([df_old, df_new], ignore_index=True)
else:
    df_combined = df_new

df_combined.to_excel(output_file, index=False)
print(f"✅ 已將 {now} 天氣資料儲存到 {output_file}")
