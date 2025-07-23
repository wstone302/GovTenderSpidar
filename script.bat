@echo off

:: 設定 Anaconda 根路徑
set CONDA_PATH=C:\Users\user\anaconda3

:: 初始化 Conda
call %CONDA_PATH%\Scripts\activate.bat

:: 啟用虛擬環境
call activate spidar

:: ✅ 改用映射磁碟機路徑
cd /d C:\Users\User\Desktop\spidar

:: 執行爬蟲腳本
:: python spidar.py
python weather_screen.py
python weather_spidar.py
python fulltext_adddatetime.py
python fulltext_spider.py


