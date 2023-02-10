from tkinter import *
from tkinter import ttk
from bs4 import BeautifulSoup
import selenium
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from tqdm import tqdm
import pandas as pd
import os

chrome_options = Options()
chrome_options.headless = True

def get_rows(driver):
    source_code = driver.page_source
    soup = BeautifulSoup(source_code, 'html.parser')
    table_body = soup.find("table", {"id": "resultTbl"})
    no_of_rows = table_body.find_all("tr").__len__()
    return no_of_rows

def getResultFromsSite():

    df1 = pd.DataFrame()

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    for i in tqdm(range(0, 20)):
        resultKey = clicked.get()

        result_id = {'B.Tech ME semester 1': 11080174, 'B.Tech ME semester 3': 11079943, 'B.Tech ME semester 5': 11079853,
                     'B.Tech ME semester 7': 11079805, 'B.Tech ECE semester 1': 11080173,'B.Tech ECE semester 5':11079947,
                     'B.Tech ECE semester 7': 11079806, ' B.Tech ECM semester 5': 11079966, 'B.Tech Civil semester 1': 11080211,
                     'B.Tech Civil semester 3':11080121,'B.Tech Civil semester 5': 11079900, 'B.Tech Civil semester 7': 11079798,
                     'B.Tech CSE semester 1' : 11080223, 'B.Tech CSE semester 3': 11080361, 'B.Tech CSE semester 5': 11079836,
                     'B.Tech CSE semester 7': 11079811,'B.Tech Civil semester 6': 11080514}
        resultId = str(result_id.get(resultKey))

        url = "https://results.pupexamination.ac.in/t8/results/results.php?rslstid=" + resultId
        driver.get(url)

        search = driver.find_element(value="inRoll")
        rollnumber = int(rollnumbervalue.get())+i
        search.send_keys(rollnumber)
        search.send_keys(Keys.ENTER)

        try:
            n_rows = get_rows(driver)
            name = driver.find_element(By.XPATH, "//table[@class ='c4']/tbody/tr/td[@class='c3']/span")
            parents = driver.find_element(By.XPATH, "//table[@class='c4']/tbody/tr[2]/td[3]/span")
            p1 = parents.text.split("\n")
            result = driver.find_element(By.XPATH, "//table[@id='resultTbl']/tbody/tr["+str(n_rows)+"]/td[3]")

            df = pd.DataFrame({"Roll Number": [rollnumber], "Name": [name.text], "Father Name": [p1[0]], "SGPA": [result.text]})
            df1 = df1.append(df)

        except Exception as e:
            pass

    print(df1)

    if(os.path.exists(f"result_{resultId + resultKey}.xlsx")):
        os.remove(f"result_{resultId + resultKey}.xlsx")

    excel = pd.ExcelWriter(f"result_{resultId + resultKey}.xlsx")
    df1.to_excel(excel, index=False)
    excel.save()

win=Tk()
win.geometry("600x600")
win.maxsize(400, 400)
win.minsize(400, 400)
win.title("Result-Finder")
win.configure(bg="light yellow")
lbl = Label(win, text="RESULT_FINDER", bg="light yellow", relief=SUNKEN)
lbl.grid(row=2, column=4)
lbl2 = Label(win, text="Select_your_Branch_and_semester", bg="light yellow")
lbl2.grid(row=8, column=3)

rollnumber = Label(win, text="Roll Number", bg="light yellow").grid(row=6, column=3)
rollnumbervalue = StringVar()
rollnumberentry = Entry(win, textvariable=rollnumbervalue)
rollnumberentry.grid(row=6, column=4)

options = {
    'B.Tech CSE semester 1', 'B.Tech CSE semester 2', 'B.Tech CSE semester 3', 'B.Tech CSE semester 4',
    'B.Tech CSE semester 5', 'B.Tech CSE semester 6', 'B.Tech CSE semester 7', 'B.Tech CSE semester 8',
    'B.Tech ME semester 1', 'B.Tech ME semester 2', 'B.Tech ME semester 3', 'B.Tech ME semester 4',
    'B.Tech ME semester 5', 'B.Tech ME semester 6', 'B.Tech ME semester 7', 'B.Tech ME semester 8',
    'B.Tech ECM semester 1', 'B.Tech ECM semester 2', 'B.Tech ECM semester 3', 'B.Tech ECM semester 4',
    'B.Tech ECM semester 5', 'B.Tech ECM semester 6', 'B.Tech ECM semester 7', 'B.Tech ECM semester 8',
    'B.Tech ECE semester 1', 'B.Tech ECE semester 2', 'B.Tech ECE semester 3', 'B.Tech ECE semester 4',
    'B.Tech ECE semester 5', 'B.Tech ECE semester 6', 'B.Tech ECE semester 7', 'B.Tech ECE semester 8',
    'B.Tech Civil semester 1', 'B.Tech Civil semester 2', 'B.Tech Civil semester 3', 'B.Tech Civil semester 4',
    'B.Tech Civil semester 5', 'B.Tech Civil semester 6', 'B.Tech Civil semester 7', 'B.Tech Civil semester 8',
    'B.Tech ME semester 1', 'B.Tech ME semester 2', 'B.Tech ME semester 3', 'B.Tech ME semester 4',
    'B.Tech ME semester 5', 'B.Tech ME semester 6', 'B.Tech ME semester 7', 'B.Tech ME semester 8',
}

def check_input(event):
    value = event.widget.get()

    if value == '':
        combo_box['values'] = options
    else:
        data = []
        for item in options:
            if value.lower() in item.lower():
                data.append(item)

        combo_box['values'] = data

clicked = StringVar()
clicked.set("")

combo_box = ttk.Combobox(win, textvariable=clicked)
combo_box['values'] = options
combo_box.bind('<KeyRelease>', check_input)
combo_box.grid(row=8, column=4)

my_button = Button(win, text="SUBMIT", command=getResultFromsSite)
my_button.grid(row=10, column=4, padx=10, pady=10)

lbl = Label(win, text="Jasleen Kaur 12001095 \n BTech Computer Science Engineering\n Punjabi University Patiala ", bg="light yellow")
lbl.grid(row=100, column=4)

win.mainloop()