import pathlib
from requests import RequestException, get
import tkinter as tk
import xlwings as xw
import pandas as pd
from datetime import datetime
from subprocess import run

OLD_FILE = "yesterday_data.csv"
NEW_FILE = "today_data.csv"
EXCEL_FILE = "Stock.xlsx"
URL = "https://www.egbroker.com/ar/MarketWatch/Index"

def get_all_data():
    today_data = get_today_data()
    if today_data.empty:
        today_data = get_today_data_no_net()
    data = pd.DataFrame.from_dict(compile_data(today_data))
    data.set_index("Ticker",inplace=True)
    data = data.sort_index()
    #apply_conditionals(data)
    workbook = xw.Book(EXCEL_FILE)
    sheet = workbook.sheets["Historic"]
    sheet["A1"].expand().options(pd.DataFrame).value = data
    historic_label.config(text="Done!")


def get_ticker_data(Ticker):
    my_entry.delete(0, tk.END)
    today_data = get_today_data()
    if today_data.empty:
        today_data = get_today_data_no_net()
    yesterday_data = get_yesterday_data()
    if not Ticker in yesterday_data.index:
        ticker_label.config(text="Ticker Not found")
        return
    data = {
        "Ticker": [Ticker],
        "Name" : [today_data.loc[Ticker]["Name"]],
        "اعلي امس" : [yesterday_data.loc[Ticker]["High"]],
        "اقل امس" : [yesterday_data.loc[Ticker]["Low"]],
        "اعلي اليوم": [today_data.loc[Ticker]["High"]],
        "اقل اليوم": [today_data.loc[Ticker]["Low"]]
    }
    append_data_to_excel(data)

def gui():
    root = tk.Tk()
    root.title("Stock")
    root.geometry("600x300")
    
    big_font = ("Comic Sans MS", 25)
    small_font = ("Comic Sans MS", 15)
    # create two frames to change the functionality of the program!
    Historic = tk.Frame(root)
    Ticker = tk.Frame(root)
    Historic.pack(fill=tk.BOTH,expand=1)
    # define the functions to change the frames back and forth.
    def change_to_historic():
        Ticker.pack_forget()
        Historic.pack(fill=tk.BOTH,expand=1)
    def change_to_Ticker():
        Historic.pack_forget()
        Ticker.pack(fill=tk.BOTH,expand=1)
    # Create labels for both frames
    tk.Label(Ticker, text="Ticker", font=big_font).pack()
    tk.Label(Historic, text="Historic", font=big_font).pack()
    
    # create buttons for switching between the two modes
    Ticker_btn = tk.Button(Historic, text="Switch to Ticker Mode",command=change_to_Ticker , font=small_font)
    Ticker_btn.pack(pady=20,padx=15,anchor=tk.N)
    
    Historic_btn = tk.Button(Ticker, text="Switch to Historic Mode",command=change_to_historic , font=small_font)
    Historic_btn.pack(pady=20,padx=15,anchor=tk.N)    
    
    # Ticker Frame Widgets
    global ticker_label
    ticker_label = tk.Label(Ticker, font=big_font,justify=tk.CENTER,text="Type a ticker")
    ticker_label.pack()
    tk.Button(Ticker,text="Refresh", font=(small_font),command=refresh).pack(pady=20,padx=15,side=tk.LEFT,anchor=tk.N)
    
    tk.Button(Ticker,text = "Submit",font = (small_font),
                          command = lambda : get_ticker_data(my_entry.get().strip().upper())).pack(pady=20,padx=15,side=tk.LEFT,anchor=tk.N)
    root.bind("<Return>",lambda event: get_ticker_data(my_entry.get().strip().upper()))

    global my_entry
    my_entry = tk.Entry(Ticker, font=small_font, justify="center")
    my_entry.pack(pady=20,side=tk.TOP)
    my_entry.focus_set()
    # Historic Widgets
    global historic_label
    historic_label =  tk.Label(Historic,text="Press the button",font=big_font)
    historic_label.pack(padx=0,pady=0)
    historic_label.lift()
    tk.Button(Historic,text="All",font=small_font,command=get_all_data, padx=25).pack(pady=20,fill=tk.X)

   

    root.mainloop()

def get_today_data():
    """
    Returns:
        Today's data: a panda dataframe of all the tickers in the URL OR AN empty dataframe if no internet
    """
    try:
        html = get(URL).content
    except RequestException:
        return pd.DataFrame()
    # uses table reading tool provided by pandas it gives back a list of tables, [1] here is the relevant table
    data = pd.read_html(html)[1]
    # drop the empty "change" column
    data.dropna(axis="columns",how="any",inplace=True)
    # change column names
    data.columns = ["Ticker","Name","Date","Last","Orders","Offer","Volume","Open","High","Low"]
    data.set_index("Ticker",inplace=True)
    return data

def get_yesterday_data():
    yesterday_dict = []
    for ticker in historic_df.index.unique():
        rows = historic_df.loc[ticker]
        last_day = rows.iloc[-1]
        try:
            high = float(last_day["High"]) 
        except IndexError:
            continue
        low = float(last_day["Low"])
        date = int(last_day["Date"])
        yesterday_dict.append({"Ticker":ticker,"High":high,"Low":low,"Date":date})
        yesterday_data = pd.DataFrame.from_dict(yesterday_dict)
        yesterday_data.set_index("Ticker",inplace=True)
    return yesterday_data

def append_data_to_excel(data):
    new_data = pd.DataFrame(data)
    new_data.set_index("Ticker",inplace=True)
    # if a file already exists we should append to it
    if pathlib.Path(EXCEL_FILE).exists():
        workbook = xw.Book(EXCEL_FILE)
        sheet = workbook.sheets["Ticker"]
        old_data = sheet["A1"].expand().options(pd.DataFrame).value
        df = pd.concat([old_data,new_data]).sort_index()
        df.drop_duplicates(subset=["Name"],keep="last", inplace=True)
        # saves the current df in the sheet!!!
        sheet["A1"].options(pd.DataFrame, header=1, index=True, expand='table').value = df
        ticker_label.config(text="Done!")
        workbook.save()
    # if the file doesnt exist we create a new one
    else:
        new_data.to_excel(EXCEL_FILE)
        # opens the excel file
        workbook = xw.Book(EXCEL_FILE)
        workbook.save()
    ticker_label.config(text="Done!")
    
def refresh():
    df = pd.read_excel(EXCEL_FILE,sheet_name="Ticker")
    df.set_index("Ticker",inplace=True)
    ticker_label.config(text="Refreshing...")
    for ticker in df.index:
        get_ticker_data(ticker)
    ticker_label.config(text="Done!")

def get_today_data_no_net():
    data = pd.read_csv(NEW_FILE)
    data.set_index("Ticker",inplace=True)
    return(data)    

def get_historic_df():
    path = "./emarket last price"
    command = f'.\ms2asc --master=MASTER --file=data.asc --indir="{path}" --ignoreName=yes --ignorePer=yes --ignoreOpenInt=yes --ignoreTime=yes --ignoreOpen=yes --printDateFrom=20220701 -q'
    run(command,shell=True)
    global historic_df
    historic_df = pd.read_csv("data.asc").dropna(axis="columns")
    historic_df.columns = ["Ticker","Date","High","Low","Close","vol"]
    historic_df.set_index("Ticker",inplace=True)
    
def compile_data(today_data):
    historic_data = []
    for ticker in today_data.index:
        if ticker in historic_df.index:
            name = today_data.loc[ticker]["Name"]
            rows = historic_df.loc[ticker]
            highest_historic_price = rows["High"].max()
            h = rows.loc[rows['High'] == highest_historic_price]["Date"][0]
            #highest_price_after_historic_high = rows.loc[(rows["Date"] >= h)]["High"].max()
            highest_historic_date = datetime.strptime(str(h), '%Y%m%d').strftime('%m/%d/%Y')
            lowest_historic_price = rows["Low"].min()
            l = rows.loc[rows['Low'] == lowest_historic_price]["Date"][0]
            highest_price_after_historic_low = float(rows.loc[(rows["Date"] >= l)]["High"].max())
            lowest_historic_date = datetime.strptime(str(l), '%Y%m%d').strftime('%m/%d/%Y')
            high_today = today_data.loc[ticker]["High"].max()
            low_today = today_data.loc[ticker]["Low"].min()
            percentage = (highest_price_after_historic_low - float(lowest_historic_price))  / float(lowest_historic_price)
            historic_data.append({"Ticker":ticker,
                                  "الاسم":name,
                                  #"تاريخ القمة التاريخية":highest_historic_date,
                                  #"قمة تاريخية":highest_historic_price,
                                  #"قمة حالية":high_today,
                                  "تاريخ القاع التاريخي": lowest_historic_date,
                                  "قاع تاريخي":lowest_historic_price,
                                  "قاع حالي":low_today,
                                  "اول قمة بعد القاع التاريخي":highest_price_after_historic_low,
                                  "النسبة بين القاع التاريخي وقمته":percentage
                                  })
    return historic_data


# pyinstaller --clean

if __name__ == "__main__":
    get_historic_df()
    get_yesterday_data()
    gui()