FINTRACK 1.3.1 a yfinance workbook constructor

TO USE

1) Download and unzip
2) python -m venv .venv 
3) .venv\scripts\activate
4) pip install -r requirements.txt  #to install dependencies
5) python main.py

ACCEPTS USER INPUT AS A STOCK TICKER RETURNS sheets\(ticker)_get_geta.xlsx 

Workbook sheets: [COMPANY INFO] [25 YEAR PRICE DATA] [CASH FLOW STATEMENT] [INCOME STATEMENT] [BALANCE SHEET]

FinTrack/
├── sheets                #Folder where new sheets are saved
├── yahooget.py           # Contains functions to fetch stock data
├── writetoworkbook.py      # Contains the function to write to Excel
├── main.py                # Contains the main program flow
└── requirements.txt   

-JBROD
