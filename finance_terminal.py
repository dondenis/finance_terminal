"""
INSTRUCTIONS:

Install required libraries using the terminal:
    pip install pandas, yfinance, tkinter, customtkinter, matplotlib 

Running the program a new window opens (dimension 1200x800). From there on proceed on working within the GUI.

The program has a search bar, in which one should enter the relevant stock ticker he is interested in. 
Upon entry, pressing the search button results in the display of the stocks 1 month trading history (in a candlechart).
If the stock ticker is invalid or nothing is entered into the search bar, an error message is displayed.

Upon seeing the relevant stocks chart, the user can adjust the timehorizon of the displayed trading history by pressing either
of the "1m", "6m", "1y", "Max" (1 month, etc.) buttons.

Furthermore, to the right of the chart, there are 5 buttons, each allowing the user to download a financial statement. 
Upon pressing either of the buttons a new window opens asking the user to select the directory where the desired file will be stored in.
The filename is the name of the company followed by the financial statement type (e.g. Tesla_Balance_Statement.xlsx).
After that the original finance terminal reopens and under the just downloaded statement button, a message is displayed confirming the successful download.

The program terminates once the window is closed.

P.S.: To be able to launch the file as a desktop application, you can convert it to an executable using pyinstaller library.

1. pip install pyinstaller
2. pyinstaller --onefile finance_terminal.py
3. select the file named "finance_terminal.exe" in the newly created "dist" folder
"""


# import relevant libraries
import datetime as dt
import pandas as pd # the scraped information from yahoo finance is in a pandas dataframe format
import os 

import yfinance as yf # yahoo finance scraping tool
import customtkinter as ctk # GUI
import tkinter as tk # GUI
from tkinter import filedialog # GUI for selecting storage dirctory for downloaded files

import matplotlib.pyplot as plt # stock chart
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg # stock chart displayed in GUI

from mplfinance.original_flavor import candlestick_ohlc # candlestick chart
import matplotlib.dates as mpl_dates # dates formatting


# Setting up window
class terminalApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        # Overall window
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.geometry("1200x800") # window size
        self.title("Denis' Finance Terminal")
        self.update()

        # Input / Search bar for the stock ticker
        ticker = tk.StringVar()
        self.symbol = ctk.CTkEntry(self, width=200, height=30, textvariable=ticker)
        self.symbol.grid(row = 0, column = 0, columnspan = 1, padx = 5, pady = 5)

        self.search_button = ctk.CTkButton(self, text="SEARCH", height=30, width=70, fg_color="gray", hover_color='green', command=self.get_ticker)
        self.search_button.grid(row = 0, column = 1, columnspan = 1, padx = 0, pady = 5)

        # Display the selected stocks name and in brackets its symbol (e.g. TSLA for Tesla Inc.)
        self.stock_label = ctk.CTkLabel(self, text="", height=20)
        self.stock_label.grid(row = 1, column = 0, padx = 5, pady = 0)

        # Options to select time frame (1 month, 6 months, 1 year, Max)
        self.button_1m = ctk.CTkButton(self, text="1m",height=5, width=50, fg_color="gray", anchor='right', hover_color='red',
                                        command=lambda: self.update_chart_timeframe("1m"))
        self.button_6m = ctk.CTkButton(self, text="6m",height=5, width=50, fg_color="gray", anchor='right', hover_color='red',
                                        command=lambda: self.update_chart_timeframe("6m"))
        self.button_1y = ctk.CTkButton(self, text="1y",height=5, width=50, fg_color="gray", anchor='right', hover_color='red',
                                        command=lambda: self.update_chart_timeframe("1y"))
        self.button_max = ctk.CTkButton(self, text="MAX",height=5, width=50, fg_color="gray", anchor='right', hover_color='red',
                                        command=lambda: self.update_chart_timeframe("max"))

        self.button_1m.grid(row = 2, column = 2, columnspan = 1, padx = 0, pady = 5)
        self.button_6m.grid(row = 2, column = 3, columnspan = 1, padx = 0, pady = 5)
        self.button_1y.grid(row = 2, column = 4, columnspan = 1, padx = 0, pady = 5)
        self.button_max.grid(row = 2, column = 5, columnspan = 1, padx = 0, pady = 5)

        # Stock chart
        self.frame = ctk.CTkFrame(master=self, height=600, width=900, fg_color="white")
        self.frame.grid(row=3, column=0, columnspan=6, rowspan=10, padx=5, pady=0)

        # Label to display potential encountered errors below chart
        self.error_label = ctk.CTkLabel(self, text="", height=20)
        self.error_label.grid(row=13, column=0, columnspan=6 ,padx=5, pady=0)

        # Buttons to download different financial statements
        self.button_balance_sheet = ctk.CTkButton(self, text="Balance Sheet",height=20, width=200, fg_color="gray", hover_color='green', command=self.on_balance_sheet)
        self.label_balance_sheet = ctk.CTkLabel(self, text="")
        
        self.button_income_statement = ctk.CTkButton(self, text="Income Statement",height=20, width=200, fg_color="gray", hover_color='green', command=self.on_income)
        self.label_income_statement = ctk.CTkLabel(self, text="")

        self.button_cash_flow_statement = ctk.CTkButton(self, text="Cash Flow Statement",height=20, width=200, fg_color="gray", hover_color='green', command=self.on_cash_flow)
        self.label_cash_flow_statement = ctk.CTkLabel(self, text="")

        self.button_dividends = ctk.CTkButton(self, text="Dividends",height=20, width=200, fg_color="gray", hover_color='green', command=self.on_dividends)
        self.label_dividends = ctk.CTkLabel(self, text="")

        self.button_shareholders = ctk.CTkButton(self, text="Shareholders",height=20, width=200, fg_color="gray", hover_color='green', command=self.on_shareholders)
        self.label_shareholders = ctk.CTkLabel(self, text="")

        self.button_balance_sheet.grid(row = 3, column = 7, padx = 50, pady = 5)
        self.label_balance_sheet.grid(row = 4, column = 7, padx = 50, pady = 5)

        self.button_income_statement.grid(row = 5, column = 7, padx = 50, pady = 5)
        self.label_income_statement.grid(row = 6, column = 7, padx = 50, pady = 5)

        self.button_cash_flow_statement.grid(row = 7, column = 7, padx = 50, pady = 5)
        self.label_cash_flow_statement.grid(row = 8, column = 7, padx = 50, pady = 5)

        self.button_dividends.grid(row = 9, column = 7, padx = 50, pady = 5)
        self.label_dividends.grid(row = 10, column = 7, padx = 50, pady = 5)

        self.button_shareholders.grid(row = 11, column = 7, padx = 50, pady = 5)
        self.label_shareholders.grid(row = 12, column = 7, padx = 50, pady = 5)


    # function called when search button is pressed (displays stock name and chart)
        
    def get_ticker(self):
        # clear chart frame everytime search button pressed to make space for new chart
        self.clear_chart_frame()
        self.error_label.configure(text='')
        self.stock_label.configure(text='')

        # reset download notifications
        self.label_balance_sheet.configure(text=f"")
        self.label_income_statement.configure(text=f"")
        self.label_cash_flow_statement.configure(text=f"")
        self.label_dividends.configure(text=f"")
        self.label_shareholders.configure(text=f"")

        # if search bar contains ticker
        if self.symbol.get().upper().strip() != "":
            stock = self.symbol.get().upper().strip()
            self.selected_stock = yf.Ticker(stock)

            try:
                self.error_label.configure(text="")
                # accessing .info attribute where the error might occur if incorrect stock ticker
                stock_name = self.selected_stock.info['shortName']
                stock_symbol = self.selected_stock.info['symbol']
                self.stock_label.configure(text=f"{stock_name} ({stock_symbol})", text_color='black')
            except:
                # display error message when entered stock ticker is incorrect
                self.stock_label.configure(text="Incorrect stock ticker.", text_color='red')
                return None

            # plot chart (candlestick)
            self.plot_candlestick()

        # if search bar is empty
        else:
            self.error_label.configure(text=f"No stock selected!", text_color='red')

    # function called to plot candlestick stock chart specifying start and enddate to display
    def plot_candlestick(self, start_date=dt.datetime.now() - dt.timedelta(days=30), end_date=dt.datetime.now()):

        if self.stock_label.cget("text") != "":
            stock = self.symbol.get().upper().strip()

            if not start_date:
                start_date = self.earliest_stock_date(stock)

            # download stock trading data for specified time horizon (format is pandas dataframe)
            data = yf.download(stock, start=start_date, end=end_date)
            data.reset_index(inplace=True)
            
            # process data
            data['Date'] = pd.to_datetime(data['Date'])
            data['Date'] = data['Date'].apply(mpl_dates.date2num)
            ohlc = data[['Date', 'Open', 'High', 'Low', 'Close']]

            # clearing existing figure
            self.clear_chart_frame()

            height = 600
            width = 900
            dpi = 100 # resolution of image
            figsize_width = width / dpi
            figsize_height = height / dpi 

            # creating new chart
            fig, ax = plt.subplots(figsize=(figsize_width, figsize_height), dpi=dpi)

            fig.patch.set_facecolor('white')  # outer background color
            ax.set_facecolor('white')  # plot area background color

            # drawing candlestick chart
            candlestick_ohlc(ax, ohlc.values, width=0.6, colorup='green', colordown='red', alpha=0.8)
            
            # defining chart layout (axis titles)
            ax.set_xlabel('Date')
            ax.set_ylabel('Price')
            date_format = mpl_dates.DateFormatter('%d-%m-%Y')
            ax.xaxis.set_major_formatter(date_format)
            fig.autofmt_xdate(which='both')
            fig.tight_layout()

            # embedding the chart into the customtkinter app
            canvas = FigureCanvasTkAgg(fig, master=self.frame)
            canvas_widget = canvas.get_tk_widget()
            canvas_widget.grid(row=3, column=0, columnspan=6, rowspan=10, padx=5, pady=0, sticky='nsew')
            canvas.draw()

        # display error message if no stock entered
        else:
            self.error_label.configure(text=f"No stock selected!", text_color='red')

    def clear_chart_frame(self):
        # clear the chart frame to make room for a new chart
        for widget in self.frame.winfo_children():
            widget.destroy()

    def update_chart_timeframe(self, timeframe):
        # update the chart based on the selected timeframe
        end_date = dt.datetime.now().date()
        if timeframe == "1m":
            start_date = end_date - dt.timedelta(days=30)
        elif timeframe == "6m":
            start_date = end_date - dt.timedelta(days=180)
        elif timeframe == "1y":
            start_date = end_date - dt.timedelta(days=365)
        elif timeframe == "max":
            start_date = None  # maximum time horizon
        self.plot_candlestick(start_date=start_date, end_date=end_date)

    # function displays window to user to select download location of file
    def choose_location(self):
        self.withdraw() # hides window
        self.file_path = filedialog.askdirectory()

    # function called upon pressing balance sheet button and performs download to excel
    def on_balance_sheet(self):
        if self.stock_label.cget("text") != "":
            balance_sheet = self.selected_stock.balance_sheet

            # Convert datetimes to timezone-naive before saving to Excel
            if isinstance(balance_sheet.index, pd.DatetimeIndex) and balance_sheet.index.tz is not None:
                balance_sheet.index = balance_sheet.index.tz_localize(None)

            self.choose_location()
            balance_sheet.to_excel(os.path.join(self.file_path, f'{self.stock_label.cget("text").split(" ", 1)[0]}_Balance_Sheet.xlsx'))
            self.deiconify() # opens window again

            # display download confirmation
            self.label_balance_sheet.configure(text=f"Balance Sheet Downloaded!", text_color='green')
        else:
            # in case no stock selected, display to user
            self.label_balance_sheet.configure(text=f"No stock selected!", text_color='red')

    # function called upon pressing income statement button and performs download to excel
    def on_income(self):
        if self.stock_label.cget("text") != "":
            income = self.selected_stock.income_stmt

            # Convert datetimes to timezone-naive before saving to Excel
            if isinstance(income.index, pd.DatetimeIndex) and income.index.tz is not None:
                income.index = income.index.tz_localize(None)

            # calls functino that allows user to select download directory
            self.choose_location()
            income.to_excel(os.path.join(self.file_path, f'{self.stock_label.cget("text").split(" ", 1)[0]}_Income_Statement.xlsx'))
            self.deiconify() # opens window again
            
            # display download confirmation
            self.label_income_statement.configure(text=f"Income Downloaded!", text_color='green')
        else:
            # in case no stock selected, display to user
            self.label_income_statement.configure(text=f"No stock selected!", text_color='red')

    # function called upon pressing cash flow statement button and performs download to excel
    def on_cash_flow(self):
        if self.stock_label.cget("text") != "":
            cash_flow = self.selected_stock.cash_flow

            # Convert datetimes to timezone-naive before saving to Excel
            if isinstance(cash_flow.index, pd.DatetimeIndex) and cash_flow.index.tz is not None:
                cash_flow.index = cash_flow.index.tz_localize(None)

            # calls functino that allows user to select download directory
            self.choose_location()
            cash_flow.to_excel(os.path.join(self.file_path, f'{self.stock_label.cget("text").split(" ", 1)[0]}_Cash_Flow.xlsx'))
            self.deiconify() # opens window again

            # display download confirmation
            self.label_cash_flow_statement.configure(text=f"Cash Flow Downloaded!", text_color='green')
        else:
            # in case no stock selected, display to user
            self.label_cash_flow_statement.configure(text=f"No stock selected!", text_color='red')

    # function called upon pressing dividends statement button and performs download to excel
    def on_dividends(self):
        if self.stock_label.cget("text") != "":
            dividends = self.selected_stock.dividends

            # Convert datetimes to timezone-naive before saving to Excel
            if isinstance(dividends.index, pd.DatetimeIndex) and dividends.index.tz is not None:
                dividends.index = dividends.index.tz_localize(None)

            # calls functino that allows user to select download directory
            self.choose_location()
            dividends.to_excel(os.path.join(self.file_path, f'{self.stock_label.cget("text").split(" ", 1)[0]}_Dividends.xlsx'))
            self.deiconify() # opens window again

            # display download confirmation
            self.label_dividends.configure(text=f"Dividends Downloaded!", text_color='green')
        else:
            # in case no stock selected, display to user
            self.label_dividends.configure(text=f"No stock selected!", text_color='red')

    # function called upon pressing shareholders button and performs download to excel
    def on_shareholders(self):
        if self.stock_label.cget("text") != "":
            shareholders = self.selected_stock.institutional_holders

            # Convert datetimes to timezone-naive before saving to Excel
            if isinstance(shareholders.index, pd.DatetimeIndex) and shareholders.index.tz is not None:
                shareholders.index = shareholders.index.tz_localize(None)

            # calls functino that allows user to select download directory
            self.choose_location()
            shareholders.to_excel(os.path.join(self.file_path, f'{self.stock_label.cget("text").split(" ", 1)[0]}_Shareholders.xlsx'))
            self.deiconify() # opens window again

            # display download confirmation
            self.label_shareholders.configure(text=f"Shareholders Downloaded!", text_color='green')
        else:
            # in case no stock selected, display to user
            self.label_shareholders.configure(text=f"No stock selected!", text_color='red')

    # function returns earliest recorded stock trading day
    def earliest_stock_date(self, stock):
        try:
            start = yf.download(stock, period="max", progress=False)
            return start.index.min()
        except Exception:
            self.error_label.configure(text=f"Error fetching data: {Exception}", text_color='red')
            return None
        
# Run window
terminal = terminalApp()
terminal.mainloop()
