# finance_terminal
The project aimed to create a desktop applica2on mimicking the func2onality of professional finance terminals like Bloomberg, coded in Python. This terminal offers financial professionals a specialized tool for accessing and analyzing stock data, alongside managing financial documents with ease and precision.

## USAGE INSTRUCTIONS:

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
