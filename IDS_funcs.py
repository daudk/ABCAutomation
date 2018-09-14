import datetime
import time
from glob import glob
from os import path, listdir
from tkinter import filedialog, END

import pandas as pd
from easygui import multpasswordbox
from numpy import abs, arange
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException

from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select



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
        if (len(fieldValues[0]) != 3):
            errmsg = errmsg + "\nSubsidiary should be a 3 letter code, e.g 'CSA', 'CUS'."
        try:
            fieldValues[1] = int(fieldValues[1])
        except:
            errmsg = errmsg + "\nYear should be a 4 digit integer, e.g 2018."
        if errmsg == "": break  # no problems found
        fieldValues = multpasswordbox(errmsg, title, fieldNames, fieldValues)

    data_340 = get_accounts_to_pull(fieldValues[2], fieldValues[3], fieldValues[0].upper() + "R0340", fieldValues[1],
                                    directory)

    process_340(data_340, directory)
    print((datetime.datetime.now() - start).total_seconds())


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

    driver = webdriver.Chrome('Resources/chromedriver.exe',chrome_options=chrome_options)
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
        j = j + 1
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


def size_accounts(accounts):
    if accounts < 1000000:
        return 0
    elif accounts < 10000000:
        return 1
    elif accounts < 50000000:
        return 2
    else:
        return 3

def parse_315_input(edit_space,subvar,yearvar,freqvar):
    accounts_str = edit_space.get(1.0,END)
    accounts = [int(i[:8]) for i in accounts_str.splitlines() if len(i)>=8]
    sub = subvar.get()
    if (sub=="CSA"):
        sub="CSAR0315"
    elif (sub=="CFS"):
        sub="CFSR0315"
    elif(sub=="CCI"):
        sub="CCIR0315"
    elif (sub=="CITS"):
        sub="CITSR315"
    else:
        sub="CUSR0315"
    print(sub)
    year = yearvar.get()
    freq=freqvar.get()

    if (freq=="Monthly"):
        freq=3
    if (freq=="Quarterly"):
        freq=2
    if (freq=="Biyearly"):
        freq=1
    if (freq=="Yearly"):
        freq=0

    run_315(accounts,sub,year,freq)

def run_315(accounts,sub,year,freq):
    print ("Pulling a total of {0} reports.".format(len(accounts*(1 if freq==0 else 2 if freq==1 else 4 if freq==2 else 12))))
    driver = webdriver.Chrome()
    driver.get("http://idssprd.cusa.canon.com/")
    driver.find_element_by_name("user").send_keys("C18889")
    driver.find_element_by_name("password").send_keys("343434Canon")
    driver.find_element_by_name("login").send_keys(Keys.RETURN)


    main_sub=True if sub=='CUSR0315' else False
    companies=True if sub=="CSAR0315" else False

    if (main_sub):
        from_path = "/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[2]/input[1]"
        to_path = "/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[2]/input[2]"
        act_from_path = "/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[2]/input[3]"
        act_to_path = "/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[2]/input[4]"
    else:
        from_path = "/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[2]/input[1]"
        to_path = "/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[2]/input[2]"
        act_from_path = "/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[2]/select[1]"
        act_to_path = "/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[2]/select[2]"

    counter = 0
    start = datetime.datetime.now()

    if (freq == 0):
        dates_start = ["01/{0}".format(year)]
        dates_end = ["12/{0}".format(year)]

    elif (freq == 1):
        dates_start = ["01/{0}".format(year), "07/{0}".format(year)]
        dates_end = ["06/{0}".format(year), "12/{0}".format(year)]

    elif (freq == 2):
        dates_start = ["01/{0}".format(year), "04/{0}".format(year), "07/{0}".format(year), "10/{0}".format(year)]
        dates_end = ["03/{0}".format(year), "06/{0}".format(year), "09/{0}".format(year), "12/{0}".format(year)]

    else:
        dates_start = ["01/{0}".format(year), "02/{0}".format(year), "03/{0}".format(year), "04/{0}".format(year) \
            , "05/{0}".format(year), "06/{0}".format(year), "07/{0}".format(year), "08/{0}".format(year) \
            , "09/{0}".format(year), "10/{0}".format(year), "11/{0}".format(year), "12/{0}".format(year)]
        dates_end = ["01/{0}".format(year), "02/{0}".format(year), "03/{0}".format(year), "04/{0}".format(year) \
            , "05/{0}".format(year), "06/{0}".format(year), "07/{0}".format(year), "08/{0}".format(year) \
            , "09/{0}".format(year), "10/{0}".format(year), "11/{0}".format(year), "12/{0}".format(year)]

    for account in accounts:
        for j in range(len(dates_start)):

            driver.get("http://idssprd.cusa.canon.com/scripts/rds/cgionline.exe?program_id=CSLMENU&rt=0&PARA=")
            top_level = driver.find_element_by_xpath('//*[@id="mnu-Main"]/li[3]')
            driver.execute_script("arguments[0].setAttribute('class','mnuopen')", top_level)
            for folder in range(1, 16):
                driver.execute_script("arguments[0].setAttribute('class','mnuopen')", \
                                      driver.find_element_by_xpath('//*[@id="mnu-FINORACLE"]/li[{0}]'.format(folder)))

            driver.find_element_by_partial_link_text(sub).click()

            for i in range(7):
                driver.find_element_by_xpath(from_path).send_keys(Keys.BACKSPACE)
            driver.find_element_by_xpath(from_path).send_keys(dates_start[j])

            for i in range(7):
                driver.find_element_by_xpath(to_path).send_keys(Keys.BACKSPACE)
            driver.find_element_by_xpath(to_path).send_keys(dates_end[j])

            driver.find_element_by_xpath(act_from_path).send_keys(str(account))
            driver.find_element_by_xpath(act_to_path).send_keys(str(account))

            if (companies):
                select = Select(driver.find_element_by_xpath('/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[4]/select[2]'))
                company_dd = Select(driver.find_element_by_xpath(
                    '/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[4]/select[4]'))
                company_dd.select_by_visible_text("311")
            else:
                select = Select(driver.find_element_by_xpath(
                    '/html/body/table[1]/tbody/tr[3]/td/form/table/tbody/tr/td[4]/select[4]'))

            select.select_by_visible_text('Payables')
            driver.find_element_by_xpath('//*[@id="submit_run"]/img').click()
            counter += 1
    time_rn = datetime.datetime.now() - start
    print(time_rn.total_seconds())