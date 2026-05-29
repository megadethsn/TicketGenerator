@echo off
setlocal
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --windowed --name TicketGenerator main.py
echo.
echo Build finished. See dist\TicketGenerator
pause
