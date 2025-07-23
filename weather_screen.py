import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from PIL import Image

# 台灣 22 個縣市（地區名稱對應 MSN）
cities = [
    "臺北市", "新北市", "桃園市", "臺中市", "臺南市", "高雄市",
    "基隆市", "新竹市", "新竹縣", "苗栗縣", "彰化縣", "南投縣",
    "雲林縣", "嘉義市", "嘉義縣", "屏東縣", "宜蘭縣", "花蓮縣",
    "臺東縣", "金門縣", "連江縣", "澎湖縣"
]

# 今日資料夾
today = datetime.now().strftime("%Y-%m-%d")
folder = f"./weather_screenshots/{today}"
os.makedirs(folder, exist_ok=True)

# Chrome 啟動設定（⚠️ 建議先不要 headless）
options = Options()
# options.add_argument("--headless")  # 建議先關掉看結果
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.5735.90 Safari/537.36")
options.add_argument("--lang=zh-TW")

# 啟動 driver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# 手動組網址，不抓首頁連結
base_url = "https://www.msn.com/zh-tw/weather/forecast/in-"

for city in cities:
    print(f"📸 擷取 {city}...")
    try:
        url = f"{base_url}{city},%20臺灣"
        driver.get(url)
        time.sleep(5)  # 等載入
        
        screenshot_path = f"{folder}/{city}.png"
        driver.save_screenshot(screenshot_path)
    except Exception as e:
        print(f"⚠️ 失敗：{city} → {e}")

driver.quit()
print("✅ 所有截圖完成")
