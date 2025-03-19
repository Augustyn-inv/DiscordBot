@echo off
set /p ticker="Enter the stock ticker (e.g., JPM): "
set /p startdate="Enter the start date (YYYY-MM-DD): "
set /p enddate="Enter the end date (YYYY-MM-DD): "

cd /d C:\Users\PATH_TO_TEMPLATE
python extract_stock_data.py %ticker% %startdate% %enddate%

pause
