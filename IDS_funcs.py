import csv
import datetime
import time
from glob import glob
from os import path, replace, listdir, makedirs
from tkinter import filedialog

import pandas as pd
from easygui import multpasswordbox
from easygui import multenterbox
from numpy import abs, arange

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

from openpyxl import load_workbook


def sample_n():
    file = filedialog.askopenfilename(initialdir = "/",title = "Select your excel file",filetypes = [("Excel Files","*.xlsx")])

    msg = "Enter Arguments to create sample"
    title = "Generating random sample"
    fieldNames = ["Worksheet index: ", "Header row: ", "Skip footer: ","Number of Samples: "]
    fieldValues = []  # we start with blanks for the values
    fieldValues = multenterbox(msg, title, fieldNames)

    extract_sample(file, fieldValues[0],fieldValues[1],fieldValues[2],fieldValues[3])

def extract_sample(file, index,header,footer,n):
    if not index:
        index = 0
    if not header:
        header=0
    if not footer:
        footer =0
    if not n:
        n=25

    df = pd.read_excel(file, sheet_name=int(index)-1,header = int(header)-1,skip_rows=int(footer))

    df2 = df.sample(int(n))

    df2.reset_index(inplace=True)
    df2['index'] = df2['index'] + 1 + int(header)
    df2 = df2.rename(columns={'index': 'Old Row Number'})

    writer = pd.ExcelWriter(file, engine='openpyxl')
    book = load_workbook(file)
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

    df2.to_excel(writer, sheet_name='New Sample', index=False)

    writer.save()



def account_pulls_from_340():
    start = datetime.datetime.now()
    directory = filedialog.askdirectory(title="Point to a new empty directory!")

    msg = "Enter Arguments to pull 340"
    title = "Requesting 340 Accounts for SOX"
    fieldNames = ["Subsidiary to pull", "Year to pull", "IDS username", "IDS password"]
    fieldValues = []  # we start with blanks for the values
    fieldValues = multpasswordbox(msg, title, fieldNames)

    # make sure that none of the fields was left blank
    while 1:
        if fieldValues == None: break
        errmsg = ""
        for i in range(len(fieldNames)):
            if fieldValues[i].strip() == "":
                errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
        if (len(fieldValues[0])!=3):
            errmsg = errmsg + "\nSubsidiary should be a 3 letter code, e.g 'CSA', 'CUS'."
        try:
            fieldValues[1]=int(fieldValues[1])
        except:
            errmsg=errmsg+"\nYear should be a 4 digit integer, e.g 2018."
        if errmsg == "": break  # no problems found
        fieldValues = multpasswordbox(errmsg, title, fieldNames, fieldValues)

    data_340 = get_accounts_to_pull(fieldValues[2], fieldValues[3], fieldValues[0].upper()+"R0340", fieldValues[1], directory)

    process_340(data_340, directory)
    print ((datetime.datetime.now()-start).total_seconds())


def download_all():
    directory = filedialog.askdirectory(title="Point to a new empty directory!") + "/"
    msg = "Download all from IDS"
    title = "IDS credentials"
    fieldNames = ["IDS username", "IDS password"]
    fieldValues = []  # we start with blanks for the values
    fieldValues = multpasswordbox(msg, title, fieldNames)

    # make sure that none of the fields was left blank
    while 1:
        if fieldValues == None: break
        errmsg = ""
        for i in range(len(fieldNames)):
            if fieldValues[i].strip() == "":
                errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
        if errmsg == "": break  # no problems found
        fieldValues = multpasswordbox(errmsg, title, fieldNames, fieldValues)

    user = fieldValues[0]
    password = fieldValues[1]

    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': directory}
    chrome_options.add_experimental_option('prefs', prefs)

    driver = webdriver.Chrome(chrome_options=chrome_options)
    driver.get("http://idssprd.cusa.canon.com/")
    driver.find_element_by_name("user").send_keys(user)
    driver.find_element_by_name("password").send_keys(password)
    driver.find_element_by_name("login").send_keys(Keys.RETURN)
    driver.get("http://idssprd.cusa.canon.com/scripts/rds/cgirpts.exe")

    start = datetime.datetime.now()
    x = arange(2, 53, 1)
    counter = 0
    try:
        while (True):
            for l in x:
                path_txt = '/html/body/form/center[1]/table/tbody/tr[2]/td[2]/table/tbody/tr[{0}]/td[4]/font'.format(
                    l)
                path_dl = '/html/body/form/center[1]/table/tbody/tr[2]/td[2]/table/tbody/tr[{0}]/td[5]/font/a[4]'.format(
                    l)
                elem = driver.find_element_by_xpath(path_txt)
                driver.find_element_by_xpath(path_dl).click()
                counter += 1
                f_name = elem.text[16:24] + '_' + elem.text[25:33] + " - " + elem.text[0:7] + " to" + elem.text[
                                                                                                      7:15]

            next_btn = driver.find_element_by_xpath(
                '/html/body/form/center[1]/table/tbody/tr[1]/td[2]/a/img').click()
    except Exception as e:
        print("All Done!")

    time_rn = datetime.datetime.now() - start
    print("Total time to download", counter, "reports was", time_rn, "seconds.")
    print("Time right now is", datetime.datetime.now())



def process_340(df_340, directory):
    df_340['size'] = df_340['amounts'].apply(size_accounts)
    df_340 = df_340[['accounts', 'Total Actuals', 'size']]
    j = 0
    while "output_{0}.csv".format(j) in listdir(directory):
        j=j+1
    df_340.to_csv(directory + '/output_{0}.csv'.format(j), index=False)



def get_accounts_to_pull(user, password, sub, year, directory):
    directory = directory + '/'

    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory': directory}
    chrome_options.add_experimental_option('prefs', prefs)

    driver = webdriver.Chrome('chromedriver.exe', chrome_options=chrome_options)
    driver.get("http://idssprd.cusa.canon.com/")
    driver.find_element_by_name("user").send_keys(user)
    driver.find_element_by_name("password").send_keys(password)
    driver.find_element_by_name("login").send_keys(Keys.RETURN)
    driver.get("http://idssprd.cusa.canon.com/scripts/rds/cgionline.exe?program_id=CSLMENU&rt=0&PARA=")

    top_level = driver.find_element_by_xpath('//*[@id="mnu-Main"]/li[3]')
    driver.execute_script("arguments[0].setAttribute('class','mnuopen')", top_level)
    for folder in range(1, 16):
        driver.execute_script("arguments[0].setAttribute('class','mnuopen')", \
                              driver.find_element_by_xpath('//*[@id="mnu-FINORACLE"]/li[{0}]'.format(folder)))

    driver.find_element_by_partial_link_text(sub).click()

    date_path = '/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[2]/input'
    for i in range(4):
        driver.find_element_by_xpath(date_path).send_keys(Keys.BACKSPACE)
    driver.find_element_by_xpath(date_path).send_keys(year)

    select = Select(
        driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[4]/select[3]'))
    select.select_by_visible_text('Payables')

    driver.find_element_by_xpath('//*[@id="submit_run"]/img').click()

    driver.get("http://idssprd.cusa.canon.com/scripts/rds/cgirpts.exe")

    path_dl = '/html/body/form/center[1]/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[5]/font/a[4]'
    path_txt = '/html/body/form/center[1]/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[3]/font'

    elem = None

    f_name = driver.find_element_by_xpath(path_txt).text[-9:-1]

    while (elem is None):
        try:
            elem = driver.find_element_by_xpath(path_dl)
        except NoSuchElementException as e:
            time.sleep(8)
            driver.get("http://idssprd.cusa.canon.com/scripts/rds/cgirpts.exe")
    elem.click()
    time.sleep(4)

    list_of_files = glob(directory + "*")
    latest_file = max(list_of_files, key=path.getctime)
    data = pd.read_excel(latest_file, header=4, skip_footer=1)
    data['accounts'] = data['Account'].str[:8]
    data['amounts'] = abs(data['Total Actuals'])

    return data


def combine_files():
    directory = filedialog.askdirectory(title="Point to directory containing files to merge:") + "/"
    counter = 0
    combined = pd.DataFrame()

    msg = "Enter parameters for File Merge"
    title = "Combine Files"
    fieldNames = ["Header Row", "Tail to skip"]
    fieldValues = []  # we start with blanks for the values
    fieldValues = multenterbox(msg, title, fieldNames)


    try:
        header_row=fieldValues[0]
        footer_rows=fieldValues[1]
    except TypeError:
        pass

    if header_row=="":
        header_row=0
    if footer_rows=="":
        footer_rows=0
    print ("header", header_row)
    print("footer", footer_rows)

    for file in listdir(directory)[:2]:
        if ".xlsx" in file:
            temp = pd.read_excel(directory + file, header=int(header_row)-1, skip_footer=int(footer_rows))
            combined = combined.append(temp, ignore_index=True)
            counter += 1
    combined.to_excel(directory+"combined_output.xlsx",index=False)
    print (counter,"files were combined and output to",directory)


def rename_all():
    directory = filedialog.askdirectory(title="Point to files to rename") + "/"

    for file in listdir(directory):
        if "0315" in file:
            temp = pd.read_excel(directory + "/" + file)
            if (len(temp.iloc[1, 0]) > 30):
                name = file[:4] + " " + temp.iloc[4, 0][15:23] + "_" + temp.iloc[5, 0][13:24] + temp.iloc[2, 0][14:16] + \
                       "-" + temp.iloc[2, 0][17:] + " to " + temp.iloc[3, 0][12:14] + "-" + temp.iloc[3, 0][15:]
            else:
                name = file[:4] + " " + temp.iloc[3, 0][15:23] + "_" + temp.iloc[4, 0][13:24] + temp.iloc[1, 0][14:16] + \
                       "-" + temp.iloc[1, 0][17:] + " to " + temp.iloc[2, 0][12:14] + "-" + temp.iloc[2, 0][15:]
            replace(directory + "/" + file, directory + "/" + name + ".xlsx")


def size_accounts(accounts):
    if accounts < 1000000:
        return 0
    elif accounts < 10000000:
        return 1
    elif accounts < 50000000:
        return 2
    else:
        return 3



def t_val_calc():
    df_summary = pd.DataFrame({'file': [], "total": []})
    min_dt = datetime.datetime(2018, 8, 1)
    counter = 0
    directory = filedialog.askdirectory(title="Point to a new empty directory!")+"/"
    output_directory = directory+"Tval_calc/"
    if not path.exists(output_directory):
        makedirs(output_directory)
    for file in listdir(directory):
        if (file[-5:] == '.xlsx' and "~$" not in file and "new" not in file and "summary" not in file):

            temp = pd.read_excel(directory + file)
            temp = temp[~temp.iloc[:, 2].isna()]
            temp.columns = temp.iloc[0]
            temp = temp.iloc[1:]
            # int_tot = np.sum(pd.to_numeric(temp['Interest'],errors='coerce'))
            int_tot = sum(pd.to_numeric(temp[temp['Date'] >= min_dt]['Interest'], errors='raise'))
            wb = load_workbook(directory + file)
            ws = wb.active
            ws['X5'] = "Total 08/01/2018 onwards:"
            ws['Z5'] = int_tot
            print(file, "    ", int_tot)
            wb.save(output_directory + file[:-5] + "_new.xlsx")
            df_summary.loc[counter] = [file, int_tot]
            counter += 1
    df_summary.to_excel(output_directory + "summary.xlsx", index=False)


def log_results():
    fields=[1,2,3,4,5,6,7,'a','b',8]
    with open("Resources/log.csv", 'a') as log:
        writer = csv.writer(log)
        writer.writerow(fields)
