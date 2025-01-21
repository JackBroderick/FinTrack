from yahooget import get_historical_data, get_shares_outstanding, income_statement_with_yfinance, balance_sheet_with_yfinance, cash_flow_statement_with_yfinance
from writetoworkbook import write_to_workbook
import time

def main():
    """
    Main program loop for fetching stock data and confirming writing to workbook.
    """
    q = 'Y'
    while q.upper() == 'Y':
        entry = input("Enter a stock ticker to fetch data for: ")

        try:
            start_time = time.perf_counter()
            print("\nFetching Outstanding shares...\n")
            hist_data = get_shares_outstanding(entry)

            # Fetch and display historical data
            print("\nFetching historical data...\n")
            hist_data = get_historical_data(entry)

            # Fetch and display income statement
            print("\nFetching income statement...\n")
            income_statement = income_statement_with_yfinance(entry)

            # Fetch and display balance sheet
            print("\nFetching balance sheet...\n")
            balance_sheet = balance_sheet_with_yfinance(entry)

            # Fetch and display cash flow statement
            print("\nFetching cash flow statement...\n")
            cash_flow = cash_flow_statement_with_yfinance(entry)
            end_time = time.perf_counter()
            print("Completed in: ", (end_time-start_time), " seconds.")


            # Ask user if they want to write the data to a new workbook
            q_write = input(f"\nWould you like to write the data for {entry} to a new workbook? (Y/N): ").strip()
            if q_write.upper() == 'Y':
                write_to_workbook(entry, hist_data, income_statement, balance_sheet, cash_flow)
            
        except KeyError:
            print("Ticker does not exist. Please try again.")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

        # Ask the user if they want to fetch data for another stock
        q = input("Would you like to enter another stock? (Y/N): ").strip()

if __name__ == '__main__':
    main()