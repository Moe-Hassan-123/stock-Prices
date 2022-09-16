from requests import RequestException, get    
import pandas as pd
import tkinter as tk
from sys import exit
from xlwings import Book
from os.path import exists

OLD_FILE = "yesterday_data.csv"
NEW_FILE = "today_data.csv"
EXCEL_FILE = "compare.xlsx"
global internet
internet = False

def main(Ticker):
    my_entry.delete(0, tk.END)
    if Ticker == "":
        result_label.config(text="Ticker was Empty")
        return
    # get data returns a dict of values of high and low
    today_data = get_livedata(Ticker)
    internet = today_data[1]
    # if there"s no internet get the last avalible live data
    values_yeseterday = get_lastdata(Ticker,OLD_FILE)
    #apply required algoraithms
    algoraithms(today_data,values_yeseterday)
    # creates a database to be used in excel
    data = {
            "Name" : [NAME],
            "Ticker" : [Ticker],
            "اعلي اليوم" : [today_data["High"]],
            "اعلي الامس" : [values_yeseterday["High"]],
            "اقل اليوم" : [today_data["Low"]],
            "اقل الامس" : [values_yeseterday["Low"]],
            "ishigher" : [ishigher]
            }
    # append the df to the excel file
    append_data(data)
    if internet == 1:
        result_label.config(text="Done!")
    else:
        result_label.config(text="Done With no internet!")

def get_livedata(Ticker):
    URL = "https://www.egbroker.com/ar/MarketWatch/Index"
    try:
        html = get(URL).content
    except RequestException:
        return get_lastdata(Ticker,NEW_FILE)
    # reads the html into a list of tables
    df_list = pd.read_html(html)
    # get the second table(the relevant one)
    df = df_list[1]
    # drop the empty "change" column
    df.dropna(axis="columns",how="any",inplace=True)
    # change column names
    df.columns = ["Ticker","Name","Date","Last","Orders","Offer","Volume","Open","High","Low"]
    df.set_index("Ticker",inplace=True)
    # save data to file
    df.to_csv("today_data.csv")
    if Ticker == "":
        return df
    if Ticker not in df.index:
        result_label.config(text="TICKER NOT FOUND")
    # gets the name from the live data to be used in excel sheet
    global NAME
    NAME = df.loc[Ticker]["Name"]
    data = df.loc[Ticker]
    # returns the data as a dict to be used in further analysis.
    return data

def get_lastdata(Ticker,File):
    df = pd.read_csv(File)
    df.set_index("Ticker" or "<Ticker>", inplace=True)
    if File == OLD_FILE:
        df.to_csv(OLD_FILE)
    if Ticker not in df.index:
        result_label.config(text="TICKER NOT FOUND")
    if File == NEW_FILE:
        global NAME
        NAME = df.loc[Ticker]["Name"]
    data = df.loc[Ticker]
    # returns the data as a dict to be used in further analysis.
    return data

def gui():
    root=tk.Tk()
    root.title("Stock")
    root.geometry("350x400")
    
    my_label = tk.Label(root, text="Ticker", font=("Helvetica",30))
    my_label.pack(pady=20)
    
    global my_entry 
    my_entry = tk.Entry(root, font=("Helvetica",20), justify="center")
    my_entry.focus_set()
    my_entry.pack(pady=20)
    
    my_button = tk.Button(root, text="Submit", font=("Helvetica",20),command=lambda : main(my_entry.get().strip().upper()))
    root.bind("<Return>",lambda event: main(my_entry.get().strip().upper()))
    my_button.pack(side="left")

    refresh_button = tk.Button(root,text="Refresh", font=("Helvetica",20),command=refresh)
    refresh_button.pack(side="right")
    
    global result_label
    result_label = tk.Label(root,text="Enter a Ticker",font=("Helvetica",25),wraplength=100)
    result_label.pack(pady=10)
    
    root.mainloop()
    return 

def append_data(data):
    # creates a new df from the data
    new_df = pd.DataFrame(data)
    # make sure its indexed correctly!
    new_df.set_index("Ticker", inplace=True)
    # if a file already exists we should append to it
    if exists("compare.xlsx"):
        workbook = Book("compare.xlsx")
        sheet = workbook.sheets[0]
        old_df = sheet["A1"].expand().options(pd.DataFrame).value
        df = pd.concat([old_df,new_df])
        df.drop_duplicates(subset=["Name"], keep="last", inplace=True)
            #sheet.used_range.value = ""
        df = df.sort_index()
        sheet.cols_right_to_left = True
        sheet.used_range.value = df
    # if the file doesnt exist we create a new one
    else:
        new_df.to_excel("compare.xlsx")
        workbook = Book("compare.xlsx")

def algoraithms(today_data,values_yeseterday):
    # algoraithm
    """
    Calculates The required algoraithm that was asked.
    all variables has to be globals so that they can be added to dataframe in main()
    """
    
    global ishigher
    if today_data["High"] > values_yeseterday["High"] and today_data["Low"] > values_yeseterday["Low"]:
        ishigher = "اعلي جديد"
    else:
        ishigher = "قاع جديد"

def refresh():
    df = pd.read_excel(EXCEL_FILE)
    df.set_index("Ticker",inplace=True)
    for ticker in df.index:
        main(ticker)

def get_historic_data():
    pass

def get_all_data():
    today_df = get_livedata("")

gui()
