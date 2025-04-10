# GovTenderSpidar

## 一、準備步驟

1. **複製專案資料夾**
   - 將整個專案資料夾（含 `run1.bat`, `script.bat`, `requirements.txt`, `spidar.py` 等檔案）**複製到桌面**。

2. **安裝 Anaconda**
   - 前往 [Anaconda 官方網站](https://www.anaconda.com/products/distribution) 下載並安裝 Anaconda（圖示為綠色蛇形圖示）。
   - 資料夾裡的 `Anaconda3-2023.07-Windows-x86_64.exe` 為 Anaconda 安裝檔，也可雙擊此檔按安裝，若已安裝可跳過此步驟。
   - 安裝完成後，即可使用 `conda` 指令進行虛擬環境與套件管理。

## 二、第一次執行

1. **建立並啟用虛擬環境**
   - **雙擊 `run1.bat`**
     - 此步驟將自動執行下列操作：
       - 使用 Python 3.11 建立名為 `spidar` 的 Conda 虛擬環境。
       - 啟用虛擬環境。
       - 安裝 `requirements.txt` 中所列之所有相依套件。

2. **執行主程式**
   - **雙擊 `script.bat`**
     - 程式將自動執行 `spidar.py` 並完成資料處理。

## 三、後續使用

- 日後使用時，僅需**雙擊 `script.bat`**，即可再次執行主程式，無需重新建立虛擬環境。

## 四、執行結果

- 程式執行後，會自動在**相同資料夾中產生 Excel 檔案作為輸出結果**。
  - 範例如：
    - `全文檢索結果.xlsx`
    - `政府採購標案彙整.xlsx`
