@echo off

set CONDA_PATH=C:\Users\User\anaconda3

call %CONDA_PATH%\condabin\conda.bat activate spidar

cd /d C:\Users\User\Desktop\spidar

python weather_screen.py || pause
python weather_spidar.py || pause
python fulltext_adddatetime.py || pause
python fulltext_spider.py || pause

pause
