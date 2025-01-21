"""
yahooget.py

This module fetches stock data from Yahoo Finance and related sources. It includes functions 
to retrieve historical data, financial statements, and other stock information.
"""

from yahoo_fin.stock_info import get_data
import yfinance as yf


def get_historical_data(ticker, start_date='1/1/2000', end_date=None):
    """
    Fetches historical stock data for the given ticker.
    """
    try:
        data = get_data(ticker, start_date=start_date, end_date=end_date)
        print("\nFetching historical data...\n")
        return data
    except ValueError as e:
        print(f"ValueError fetching historical data for {ticker}: {e}")
        return None
    except Exception as e:
        print(f"Error fetching historical data for {ticker}: {e}")
        return None


def get_shares_outstanding(ticker):
    """
    Fetches the number of shares outstanding for the given ticker.
    """
    try:
        stock = yf.Ticker(ticker)
        outstanding = stock.info.get("sharesOutstanding")
        return outstanding
    except KeyError as e:
        print(f"KeyError fetching shares outstanding for {ticker}: {e}")
        return None
    except Exception as e:
        print(f"Error fetching shares outstanding for {ticker}: {e}")
        return None


def income_statement_with_yfinance(ticker):
    """
    Fetches the income statement using yfinance.
    """
    try:
        stock = yf.Ticker(ticker)
        financials = stock.financials
        if financials.empty:
            print(f"No income statement data available for {ticker}.")
            return None
        print("\nFetching income statement...\n")
        return financials
    except KeyError as e:
        print(f"KeyError fetching income statement for {ticker}: {e}")
        return None
    except Exception as e:
        print(f"Error fetching income statement for {ticker}: {e}")
        return None


def balance_sheet_with_yfinance(ticker):
    """
    Fetches the balance sheet using yfinance.
    """
    try:
        stock = yf.Ticker(ticker)
        balance_sheet = stock.balance_sheet
        if balance_sheet.empty:
            print(f"No balance sheet data available for {ticker}.")
            return None
        print("\nFetching balance sheet...\n")
        return balance_sheet
    except KeyError as e:
        print(f"KeyError fetching balance sheet for {ticker}: {e}")
        return None
    except Exception as e:
        print(f"Error fetching balance sheet for {ticker}: {e}")
        return None


def cash_flow_statement_with_yfinance(ticker):
    """
    Fetches the cash flow statement using yfinance.
    """
    try:
        stock = yf.Ticker(ticker)
        cashflow = stock.cashflow
        if cashflow.empty:
            print(f"No cash flow statement data available for {ticker}.")
            return None
        print("\nFetching cash flow statement...\n")
        return cashflow
    except KeyError as e:
        print(f"KeyError fetching cash flow statement for {ticker}: {e}")
        return None
    except Exception as e:
        print(f"Error fetching cash flow statement for {ticker}: {e}")
        return None


def get_stock_info(ticker):
    """
    Fetches and returns selected stock information as a dictionary.
    """
    try:
        stock = yf.Ticker(ticker)
        info = stock.info

        selected_info = {
            "Company Name": info.get("longName"),
            "Symbol": info.get("symbol"),
            "Sector": info.get("sector"),
            "Industry": info.get("industry"),
            "Employees": info.get("fullTimeEmployees"),
            "Country": info.get("country"),
            "Address": info.get("address1"),
            "City": info.get("city"),
            "State": info.get("state"),
            "Zip": info.get("zip"),
            "Phone": info.get("phone"),
            "Website": info.get("website"),
            "Exchange": info.get("exchange"),
            "Market Cap": info.get("marketCap"),
            "Shares Outstanding": get_shares_outstanding(ticker),
            "P/E Ratio": info.get("forwardPE"),
            "Beta": info.get("beta"),
            "Previous Close": info.get("previousClose"),
            "Regular Market Day High": info.get("regularMarketDayHigh"),
            "Regular Market Day Low": info.get("regularMarketDayLow"),
            "Fifty-Two Week High": info.get("fiftyTwoWeekHigh"),
            "Fifty-Two Week Low": info.get("fiftyTwoWeekLow"),
            "Dividend Rate": info.get("dividendRate"),
            "Dividend Yield": info.get("dividendYield"),
        }
        return selected_info
    except KeyError as e:
        print(f"KeyError fetching stock info for {ticker}: {e}")
        return None
    except Exception as e:
        print(f"An error occurred while fetching data for {ticker}: {e}")
        return None


def main():
    """Main entry point of the script."""
    pass


if __name__ == "__main__":
    main()
