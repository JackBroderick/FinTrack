import os
from yahoo_fin.stock_info import get_data, tickers_sp500, tickers_nasdaq, tickers_other, get_quote_table, get_income_statement
import yfinance as yf
import time
from writetoworkbook import write_to_workbook  # Import the write_to_workbook function

def get_historical_data(ticker, start_date='1/1/2000', end_date=None):
    """
    Fetches historical stock data for the given ticker and prints it.
    """
    try:
        # Fetch historical data
        datamess = get_data(ticker, start_date=start_date, end_date=end_date)
        print("\nHistorical Data:")
        print(datamess)
        return datamess
    except Exception as e:
        print(f"Error fetching historical data for {ticker}: {e}")

def get_shares_outstanding(ticker):
    try:
        # Fetch the ticker data
        stock = yf.Ticker(ticker)
        # Get shares outstanding
        outstanding = stock.info.get("sharesOutstanding")
        
        if outstanding is None:
            print(f"No shares data available for {ticker}.")
        else:
            print(f"\nShares outstanding for {ticker}: {outstanding}")
        
        return outstanding
    except Exception as e:
        print(f"Error fetching shares outstanding for {ticker} with yfinance: {e}")

def income_statement_with_yfinance(ticker):
    """
    Fetches the income statement using the yfinance library and prints it.
    """
    try:
        stock = yf.Ticker(ticker)
        financials = stock.financials  # Income statement
        if financials.empty:
            print(f"No income statement data available for {ticker}.")
        else:
            print(f"\nIncome Statement for {ticker}:")
            print(financials)
        return financials
    except Exception as e:
        print(f"Error fetching income statement for {ticker} with yfinance: {e}")

def balance_sheet_with_yfinance(ticker):
    """
    Fetches the balance sheet using the yfinance library and prints it.
    """
    try:
        stock = yf.Ticker(ticker)
        balance_sheet = stock.balance_sheet  # Balance sheet
        if balance_sheet.empty:
            print(f"No balance sheet data available for {ticker}.")
        else:
            print(f"\nBalance Sheet for {ticker}:")
            print(balance_sheet)

        return balance_sheet
    except Exception as e:
        print(f"Error fetching balance sheet for {ticker} with yfinance: {e}")

def cash_flow_statement_with_yfinance(ticker):
    """
    Fetches the cash flow statement using the yfinance library and prints it.
    """
    try:
        stock = yf.Ticker(ticker)
        cashflow = stock.cashflow  # Cash flow statement
        if cashflow.empty:
            print(f"No cash flow statement data available for {ticker}.")
        else:
            print(f"\nCash Flow Statement for {ticker}:")
            print(cashflow)
        return cashflow
    except Exception as e:
        print(f"Error fetching cash flow statement for {ticker} with yfinance: {e}")

def main():
    pass

if __name__ == '__main__':
    main()

