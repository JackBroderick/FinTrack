FINTRACK 1.3.2 - yfinance excel workbook constructor

***TO USE***

1) Download and unzip
2) cd FinTrack
3) python -m venv .venv 
4) .venv\scripts\activate
5) pip install -r requirements.txt 
6) python main.py

WILL ACCEPT A STOCK TICKER RETURN sheets\(ticker)_stock_data.xlsx 

Workbook sheets: [COMPANY INFO] [25 YEAR PRICE DATA] [CASH FLOW STATEMENT] [INCOME STATEMENT] [BALANCE SHEET]


FinTrack/
├── sheets                #Folder where new sheets are saved
├── yahooget.py           # Contains functions to fetch stock data
├── writetoworkbook.py      # Contains the functions to write to Excel
├── main.py                # Contains the main program flow
└── requirements.txt   

-JBROD
