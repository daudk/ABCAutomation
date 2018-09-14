import time
import datetime
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

def s21_invoices():
    try:
        import pywinauto as pywin
        start = datetime.datetime.now()
        driver = webdriver.Ie("Resources/IEDriverServer.exe")
        driver.set_window_size(1124, 850)
        driver.get("https://s21.cusa.canon.com/wholesale/main.jsp")

        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.NAME, "wfSubDocumentId")))

        driver.execute_script("JumpToOtherBusinessApp('C1')")

        counter = 0

        inv_nos = [71159881, 71173008, 71220857, 71263759, 71273786, 71286753, 71312994, 71335814, 71360486, 71374933,
                   71387713, 71433663]

        for inv in inv_nos:

            WebDriverWait(driver, 25).until(
                EC.presence_of_element_located((By.NAME, 'invNumSrchTxt_H')))

            driver.find_element_by_name('invNumSrchTxt_H').send_keys(str(inv))
            driver.find_element_by_name("invDt_B").clear()
            driver.find_element_by_name("invDt_A").clear()

            s3 = driver.find_element_by_name("Search_Invoice")
            s3.send_keys("")
            s3.click()

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "xxChkBox")))

            s2 = driver.find_element_by_name("xxChkBox")
            s2.send_keys("")
            s2.click()

            s1 = driver.find_element_by_name("Print_Invoice")
            s1.click()
            f1 = False
            while (f1 == False):
                try:
                    pdf = pywin.Application().connect(title_re="https://s21.cusa.canon.com/ - S21 - Internet Explorer")
                    f1 = True
                except:
                    continue

            pdf = pdf.top_window()
            static = pdf.Static
            static.SetFocus()
            time.sleep(1)
            static.TypeKeys("+^S")

            save_app = pywin.application.Application()

            f2 = False
            while (not f2):
                try:
                    save_dialog = save_app.connect(title_re=".*Save a Copy*")
                    f2 = True
                except:
                    continue
            save_dialog_top = save_dialog.top_window()

            #     address=save_dialog_top['Toolbar3']
            #     address.Click()
            #     address.TypeKeys(dl_path)

            #     name_bar = save_dialog_top['5']
            #     name_bar.DoubleClick()

            save_dialog_top.TypeKeys(str(inv).replace("/", "-"))

            final = save_dialog_top['&Save']

            final.ClickInput()

            counter += 1

            time.sleep(1)
            pdf.close()

            actions = ActionChains(driver)
            actions.send_keys(Keys.F10)
            actions.perform()

        # for i in os.listdir(dl_path):
        #     os.rename(dl_path+i,dl_path+i[:-6]+i[-4:])

        print("Total Reports: ", counter)
        print("Total Time spend: ", (datetime.datetime.now() - start).total_seconds())
    except:
        del pywin