from yahooget import (
    get_historical_data,
    get_shares_outstanding,
    income_statement_with_yfinance,
    balance_sheet_with_yfinance,
    cash_flow_statement_with_yfinance,
)
from writetoworkbook import write_to_workbook
import time


def main():
    """
    Main program loop for fetching stock data and confirming writing to a workbook.
    """
    user_input = 'Y'
    while user_input.upper() == 'Y':
        entry = input("Enter a stock ticker to fetch data for: ").strip()

        try:
            start_time = time.perf_counter()

            # Fetch historical data
            hist_data = get_historical_data(entry)

            # Fetch income statement
            income_statement = income_statement_with_yfinance(entry)

            # Fetch balance sheet
            balance_sheet = balance_sheet_with_yfinance(entry)

            # Fetch cash flow statement
            cash_flow = cash_flow_statement_with_yfinance(entry)

            end_time = time.perf_counter()
            print(f"\nData fetching completed in: {end_time - start_time:.2f} seconds.")

            # Ask the user if they want to write the data to a workbook
            write_input = input(
                f"\nWould you like to write the data for {entry} to a new workbook? (Y/N): "
            ).strip()

            if write_input.upper() == 'Y':
                write_to_workbook(entry, hist_data, income_statement, balance_sheet, cash_flow)

        except KeyError:
            print("Ticker does not exist. Please try again.")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

        # Ask the user if they want to fetch data for another stock
        user_input = input("Would you like to enter another stock? (Y/N): ").strip()


if __name__ == '__main__':
    main()
