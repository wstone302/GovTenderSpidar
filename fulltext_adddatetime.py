from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from datetime import datetime, timedelta
from collections import Counter

# 關鍵字與對應產業分類
keywords = [
    "測繪業", "測量", "測繪", "檢測", "地形", "地籍圖", "地籍", "套疊", "套繪", "圖資", "水深", "淤積", "海床",
    "空拍", "航拍", "無人機", "飛行載具", "UAV", "UAS", "圖根點", "監測", "斷面", "掃描", "點雲",
    "LIDAR", "BIM", "光達", "沉陷", "水庫", "雷射", "疏濬", "疏浚", "地形收方", "排水", "電塔", "湖", "海洋", "海岸",
    "多音束", "單音束", "SBES", "離岸風電", "透地雷達", "聲納", "數值地形", "建模", "閥栓", "計畫樁清理", "補充調查",
    "水資源分署", "大觀發電廠", "萬大發電廠", "河川分署", "技師事務所", "精密儀器批發業"
]

# 定義近一個月範圍
today = datetime.today()
one_month_ago = today - timedelta(days=60)

driver = webdriver.Chrome()
wait = WebDriverWait(driver, 15)
all_data = []

for keyword in keywords:
    category_name = keyword.strip()
    print(f"\n🔍 查詢全文關鍵字：{category_name}")
    driver.get("https://web.pcc.gov.tw/prkms/tender/common/bulletion/indexBulletion")

    try:
        input_box = wait.until(EC.presence_of_element_located((By.NAME, "querySentence")))
        input_box.clear()
        input_box.send_keys(category_name)
        driver.find_element(By.XPATH, "//a[@title='查詢']").click()
        time.sleep(2)

        page = 1
        break_outer_loop = False  # 每個關鍵字初始化為不跳出

        while True:
            wait.until(EC.visibility_of_element_located((By.XPATH, "//table[@id='bulletion']/tbody/tr")))
            rows = driver.find_elements(By.XPATH, "//table[@id='bulletion']/tbody/tr")
            print(f"✅ 第 {page} 頁，共 {len(rows)} 筆")

            for idx, row in enumerate(rows):
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 6:
                    try:
                        category = cols[1].text.strip()
                        agency = cols[2].text.strip()
                        tender_block = cols[3].text.strip().split("\n")
                        tender_id = tender_block[0].strip() if len(tender_block) > 0 else ""
                        tender_name = tender_block[1].strip() if len(tender_block) > 1 else ""
                        announce_date = cols[4].text.strip()
                        deadline_date = cols[6].text.strip()
                        detail_link = cols[-1].find_element(By.TAG_NAME, "a").get_attribute("href")

                        # 日期轉換與近一個月篩選
                        try:
                            parts = announce_date.split("/")
                            if len(parts) == 3:
                                year = int(parts[0]) + 1911
                                month = int(parts[1])
                                day = int(parts[2])
                                announce_dt = datetime(year, month, day)
                            else:
                                raise ValueError("格式錯誤")

                            if announce_dt < one_month_ago:
                                print(f"🛑 發現超過一個月前的日期：{announce_date}，停止翻頁")
                                break_outer_loop = True
                                break  # 終止 row 迴圈

                        except Exception as e:
                            print(f"❗ 無法解析公告日期：{announce_date}，略過此筆")
                            continue

                        item = {
                            "來源關鍵字": category_name,
                            "種類": category,
                            "機關名稱": agency,
                            "標案案號": tender_id,
                            "標案名稱": tender_name,
                            "招標公告日期": announce_date,
                            "截止投標日期": deadline_date,
                            "詳細頁連結": detail_link
                        }

                        print(f"✅ 擷取成功：{item}")
                        all_data.append(item)

                    except Exception as e:
                        print(f"❗ 欄位解析失敗：{e}")
                        continue

            if break_outer_loop:
                break  # 終止 while 翻頁迴圈

            next_button = driver.find_elements(By.XPATH, "//a[contains(text(), '下一頁')]")
            if next_button:
                print(f"➡ 進入第 {page + 1} 頁")
                next_button[0].click()
                page += 1
                time.sleep(2)
            else:
                print("📄 已無下一頁")
                break

    except Exception as e:
        print(f"❌ 查詢失敗：{category_name}，錯誤：{e}")
        continue

driver.quit()

# ✅ 匯出 Excel（指定欄位順序）
output_path = r"C:\Users\User\Desktop\spidar\全文檢索結果_datetime.xlsx"
df = pd.DataFrame(all_data)
df.drop_duplicates(subset=["標案名稱", "詳細頁連結"], keep='first', inplace=True)

columns_order = [
    "來源關鍵字", "種類", "機關名稱", "標案案號", "標案名稱",
    "招標公告日期", "截止投標日期", "詳細頁連結"
]
df = df[columns_order]
df.to_excel(output_path, index=False)
print(f"\n📦 最終擷取筆數：{len(df)}\n✅ 已儲存至：{output_path}")

# 顯示關鍵字統計
keyword_counter = Counter([d['來源關鍵字'] for d in all_data])
print("\n📊 擷取筆數統計：")
for k, v in keyword_counter.items():
    print(f"- {k}: {v} 筆")
