@echo off
echo [1/4] Creating conda environment 'spidar' with Python 3.11...
conda create -n spidar python=3.11 -y

echo [2/4] Activating environment...
call conda activate spidar

echo [3/4] Changing directory to project folder...
cd /d C:\Users\user\Desktop\spidar

echo [4/4] Installing dependencies from requirements.txt...
pip install -r requirements.txt

echo Setup complete. Environment 'spidar' is ready.
pause
