@echo off

:: 設定 Anaconda 根路徑
set CONDA_PATH=C:\Users\user\anaconda3

:: 初始化 Conda
call %CONDA_PATH%\Scripts\activate.bat

:: ✅ 此處改為 call activate spidar，而非直接 conda activate（比較穩）
call activate spidar

:: 切換到爬蟲腳本所在資料夾
cd /d C:\Users\user\Desktop\spidar

:: 執行爬蟲
python spidar.py
python fulltext_spider.py

:: 暫停畫面供除錯用
pause
