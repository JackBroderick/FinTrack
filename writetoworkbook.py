import os
from yahoo_fin.stock_info import get_data, tickers_sp500, tickers_nasdaq, tickers_other, get_quote_table, get_income_statement
import yfinance as yf
import openpyxl

def adjust_column_width(sheet):
    """
    Adjusts the column width of all columns in the sheet based on the longest string in each column.
    """
    for column in sheet.columns:
        max_length = 0
        column_name = column[0].column_letter  # Get the column name (e.g. 'A', 'B', 'C', ...)
        
        for cell in column:
            try:
                # Get the length of the string representation of each cell and keep track of the max length
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        # Adjust the column width, add extra space for padding
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column_name].width = adjusted_width

def write_to_workbook(ticker, hist_data, income_statement, balance_sheet, cash_flow):
    """
    Writes the fetched data to a new Excel workbook, including all years for financial statements 
    and a column for the average percent change across the years as a decimal.
    """
    file_path = f"sheets/{ticker}_stock_data.xlsx"
    
    # Check if the file already exists
    if os.path.exists(file_path):
        overwrite = input(f"The file {ticker}_stock_data.xlsx already exists. Would you like to overwrite it? (Y/N): ").strip()
        if overwrite.upper() != 'Y':
            print(f"Skipping the overwrite. Data for {ticker} will not be saved.")
            return

    # Create a new workbook
    wb = openpyxl.Workbook()

    # Write historical data
    sheet = wb.create_sheet(title=f"{ticker} - Historical Data")
    sheet.append(["Date", "Open", "High", "Low", "Close", "Volume"])
    for index, row in hist_data.iterrows():
        sheet.append([index.date(), row["open"], row["high"], row["low"], row["close"], row["volume"]])

    # Adjust column widths for historical data sheet
    adjust_column_width(sheet)

    # Write financial statements with average percent change
    for name, data in [("Income Statement", income_statement), 
                       ("Balance Sheet", balance_sheet), 
                       ("Cash Flow Statement", cash_flow)]:
        sheet = wb.create_sheet(title=f"{ticker} - {name}")
        
        # Write column headers
        headers = ["Metric"] + list(data.columns) + ["Avg Change (Decimal)"]
        sheet.append(headers)

        for row_label in data.index:
            # Get values for all years
            values = data.loc[row_label].tolist()

            try:
                # Calculate percent changes and average percent change
                percent_changes = data.loc[row_label].pct_change(fill_method=None).dropna()
                avg_change = percent_changes.mean() if not percent_changes.empty else None
            except ZeroDivisionError:
                avg_change = None  # In case of zero division error, set as None

            # Append the row
            sheet.append([row_label] + values + [avg_change])

        # Adjust column widths
        adjust_column_width(sheet)

    # Save the workbook
    wb.save(file_path)
    print(f"\nData for {ticker} has been written to {file_path}")

    def main():
        pass
    
    if __name__ == '__main__':
        main()
        
