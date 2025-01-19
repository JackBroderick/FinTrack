import os
import openpyxl
from openpyxl.styles import NamedStyle, Font
from yahooget import get_stock_info  # Import get_stock_info from yahooget.py

# Create dollar and percent styles
def create_dollar_style():
    dollar_style = NamedStyle(name="dollar_style", number_format='"$"#,##0.00')
    return dollar_style

def create_percent_style():
    percent_style = NamedStyle(name="percent_style", number_format="0.0%")
    return percent_style

def adjust_column_width(sheet):
    for column in sheet.columns:
        max_length = 0
        column_name = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column_name].width = adjusted_width

def write_to_workbook(ticker, hist_data, income_statement, balance_sheet, cash_flow):
    os.makedirs('sheets', exist_ok=True)
    
    file_path = f"sheets/{ticker}_stock_data.xlsx"
    
    if os.path.exists(file_path):
        overwrite = input(f"The file {ticker}_stock_data.xlsx already exists. Would you like to overwrite it? (Y/N): ").strip()
        if overwrite.upper() != 'Y':
            print(f"Skipping the overwrite. Data for {ticker} will not be saved.")
            return

    wb = openpyxl.Workbook()

    # Remove the default "Sheet" tab if it exists
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    
    # Fetch stock info
    stock_info = get_stock_info(ticker)
    
    # Create and populate "Stock Info" sheet with the data fetched from get_stock_info
    sheet = wb.create_sheet(title=f"{ticker} - Stock Info")
    sheet.append(["Metric", "Value"])

    # Write the stock information to the "Stock Info" sheet
    for key, value in stock_info.items():
        sheet.append([key, value])

    adjust_column_width(sheet)

    # Historical Data
    sheet = wb.create_sheet(title=f"{ticker} - Historical Data")
    sheet.append(["Date", "Open", "High", "Low", "Close", "Volume"])
    for index, row in hist_data.iloc[::-1].iterrows():
        sheet.append([index.date(), row["open"], row["high"], row["low"], row["close"], row["volume"]])

    # Apply dollar formatting to numerical columns only (Open, High, Low, Close, Volume)
    for row in sheet.iter_rows(min_row=2, min_col=2, max_col=6):  # Starting from the second row
        for cell in row:
            if isinstance(cell.value, (int, float)):  # Check if the cell contains a numeric value
                cell.number_format = '"$"#,##0.00'  # Apply dollar format directly

    adjust_column_width(sheet)

    # Financial Statements
    for name, data in [("Income Statement", income_statement), 
                       ("Balance Sheet", balance_sheet), 
                       ("Cash Flow Statement", cash_flow)]:
        sheet = wb.create_sheet(title=f"{ticker} - {name}")
        reversed_data = data.iloc[:, ::-1]
        headers = ["Metric"] + list(reversed_data.columns) + ["Avg Yearly Change"]
        sheet.append(headers)

        for row_label in reversed_data.index:
            values = reversed_data.loc[row_label].tolist()
            try:
                percent_changes = reversed_data.loc[row_label].pct_change(fill_method=None).dropna()
                avg_change = percent_changes.mean() if not percent_changes.empty else None
            except ZeroDivisionError:
                avg_change = None

            sheet.append([row_label] + values + [avg_change])

            # Apply dollar formatting to numeric columns
            for col in range(1, len(values) + 1):  # Iterate over the value columns
                if isinstance(values[col - 1], (int, float)):  # Ensure it's a number
                    sheet.cell(row=sheet.max_row, column=col+1).number_format = '"$"#,##0.00'  # Dollar format

            # Apply percentage formatting for average change
            avg_change_cell = sheet.cell(row=sheet.max_row, column=len(values) + 2)
            if avg_change is not None:
                avg_change_cell.value = avg_change
                avg_change_cell.number_format = "0.0%"  
                
                if avg_change > 0:
                    avg_change_cell.font = Font(color="228B22")
                elif avg_change < 0:
                    avg_change_cell.font = Font(color="FF0000")

        adjust_column_width(sheet)

    wb.save(file_path)
    print(f"Data for {ticker} has been written to {file_path}")



