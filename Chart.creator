import yfinance as yf
import openpyxl
import sys
import shutil
from datetime import datetime
import os
from openpyxl.styles import Font


# Function to fetch stock data and update Excel
def fetch_and_update_stock_data(ticker, start_date, end_date, template_file, new_file):
    # Download stock data
    from datetime import datetime, timedelta

    # Convert end_date to a datetime object, add one day, and format it back to a string
    end_date_extended = (datetime.strptime(end_date, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")

    # Use the extended end_date in the download
    stock_data = yf.download(ticker, start=start_date, end=end_date_extended)

    # Create a duplicate of the template file
    shutil.copy(template_file, new_file)

    # Open the newly created Excel workbook and sheet
    wb = openpyxl.load_workbook(new_file)
    sheet = wb.active  # Assuming you're working with the active sheet

    # Fill values
    sheet['C2'] = ticker
    sheet['C3'] = start_date
    sheet['C4'] = end_date

    # Hyperlink for ticker
    url = f'https://finance.yahoo.com/quote/{ticker}'
    sheet['C5'] = f'=HYPERLINK("{url}", "{url}")'

    # Initialize the maximum and minimum stock price variables
    max_price = float('-inf')
    min_price = float('inf')

    # Define the font style for new rows (Aptos Narrow, size 12)
    font_style = Font(name='Aptos Narrow', size=12)

    # Loop to fill the stock data (dates and prices) starting from row 7
    current_row = 7  # Start from row 7 for stock data
    rows_inserted = False  # Flag to track if any rows are inserted

    for date, row in stock_data.iterrows():
        if current_row > 27:
            # Insert a new row for each new price if we exceed row 27
            sheet.insert_rows(current_row)
            rows_inserted = True

        # Format and insert the date in column B
        sheet[f'B{current_row}'] = date.strftime('%d-%m-%Y')

        # Ensure the 'Close' price is prefixed with "=" and formatted with two decimal places
        close_price = round(float(row['Close'].iloc[0]), 2)  # Fix for FutureWarning
        sheet[f'C{current_row}'] = f"={close_price:.2f}"  # Add "=" to treat as a number in Excel with two decimals

        # Apply font style to new row
        sheet[f'B{current_row}'].font = font_style
        sheet[f'C{current_row}'].font = font_style

        # Update max and min prices for later insertion
        if close_price > max_price:
            max_price = close_price
        if close_price < min_price:
            min_price = close_price

        # Move to the next row for the next price
        current_row += 1

    # Determine the row for the formula
    if not rows_inserted:
        # If no rows were inserted, place the formula in C29
        formula_row = 29
        sheet['C29'] = f"=(C{current_row - 1} - C7) / C7"
    else:
        # If rows were inserted, place the formula in the row after the last inserted price (+1)
        formula_row = current_row + 1  # The row after the last price
        last_price_cell = f'C{current_row - 1}'  # The last entered price will be in the last row of column C
        sheet[f'C{formula_row}'] = f"=({last_price_cell} - C7) / C7"  # Update the formula reference dynamically

    # Insert minimum price and maximum price two rows after the formula
    if rows_inserted:
        # If rows were inserted, place min and max values after the formula
        min_price_row = formula_row + 2  # Two rows after the formula
        max_price_row = min_price_row + 1  # One row after the min value
    else:
        # If no rows were inserted, place min and max values two rows after the formula
        min_price_row = formula_row + 2  # Two rows after the formula
        max_price_row = min_price_row + 1  # One row after the min value

    # Insert min and max prices prefixed with "=" and formatted with two decimal places
    sheet[f'C{min_price_row}'] = f"={min_price:.2f}"  # Minimum price as a number with two decimals
    sheet[f'C{max_price_row}'] = f"={max_price:.2f}"  # Maximum price as a number with two decimals

    # Apply font style to the new rows (min and max price rows)
    sheet[f'C{min_price_row}'].font = font_style
    sheet[f'C{max_price_row}'].font = font_style

    # Save the updated workbook
    wb.save(new_file)
    print(f"Excel file created and filled with stock data for {ticker}.")


# Main
if __name__ == "__main__":
    # Get the command line arguments
    if len(sys.argv) != 4:
        print("Usage: python extract_stock_data.py <ticker> <start_date> <end_date>")
        sys.exit(1)

    ticker = sys.argv[1]
    start_date = sys.argv[2]
    end_date = sys.argv[3]

    # Paths for template and new file
    template_file = r"C:\Users\mochm\Desktop\Studies\SIF\Straipsniai\Template\sifTemplate.xlsx"

    # Create the new file name with ticker and "_graph"
    new_file_name = f"{ticker}_graph.xlsx"
    new_file = os.path.join(r"C:\Users\mochm\Desktop", new_file_name)  # Saving the new file on the Desktop

    # Call the function to fetch data and update the duplicate Excel file
    fetch_and_update_stock_data(ticker, start_date, end_date, template_file, new_file)
