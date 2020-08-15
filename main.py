#! python3

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import time
import datetime
import pyautogui
from PIL import Image

bmp_print = Image.open('drukuj.bmp')

binary = FirefoxBinary("C:\\Program Files\\Mozilla Firefox\\firefox.exe")
browser = webdriver.Firefox(firefox_binary=binary, executable_path="F:\\python38\\geckodriver.exe")

login = 'login'
password = 'password'
txt_filename = time.strftime("%Y%m%d_%H%M")
log_file = open('base_' + txt_filename + '.txt', 'a')


def scanning():
    global accession_numbers
    global auto_increment

    accession_numbers = []
    # input/scan accession number
    accession_number = input(
        'Scan accession number or press D for printing.\npress Q to quit\n').lower()

    while accession_number != 'd' or accession_number != 'q' or len(accession_number) == 0:
        try:
            # force accession_number to be an integer
            accession_number = int(accession_number)
            accession_numbers.append(accession_number)
            if auto_increment == 'yes':
                print('input accession number or press D for printing.')
                # auto_increment and  remove "None" by splitting increment and input into 2 lines
                pyautogui.typewrite(str(int(accession_number) + 1))
                accession_number = input().lower()
            else:
                accession_number = input('Scan accession number or press D for printing.\n').lower()
        except ValueError:
            if accession_number == 'd':
                break
            elif accession_number == 'q':
                auto_increment = 'q'
                break
            else:
                print('\nError.\nScan accession number or press D for printing.')
                # let the user correct wrong value
                pyautogui.typewrite(str(accession_number))
                accession_number = input().lower()
    if accession_number == 'q':
        auto_increment = 'q'


def print_barcodes(accession_numbers_list):
    global auto_increment

    for single_number in accession_numbers:

        # wait for fully loaded site
        delay = 5
        try:
            my_elem = WebDriverWait(browser, delay).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#ctl00_frmRejestrIntrodukcji")))
            print("Page is ready!")
            frame = browser.find_element_by_css_selector("#ctl00_frmRejestrIntrodukcji")
            browser.switch_to.frame(frame)
            try:
                i_elem = WebDriverWait(browser, delay).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#RejestrIntrodukcji1_JXPGrid_RI_body")))
                print("Subpage is ready!")
            except TimeoutException:
                print("Too long for iframe presence")
        except TimeoutException:
            print("Too long")

        # find accession number
        browser.switch_to.default_content()
        browser.find_element_by_css_selector("#ctl00_RibbonGroupIntrodukcja_rtd_main_ri_numer_akcesyjny_text").clear()
        browser.find_element_by_css_selector(
            "#ctl00_RibbonGroupIntrodukcja_rtd_main_ri_numer_akcesyjny_text").send_keys(single_number)
        time.sleep(1.5)
        browser.find_element_by_css_selector(
            "#Table4 > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2) > img:nth-child(1)").click()

        # check if accession number exist
        try:
            # highlight table row
            time.sleep(2)
            frame = browser.find_element_by_css_selector("#ctl00_frmRejestrIntrodukcji")
            browser.switch_to.frame(frame)
            browser.find_element_by_css_selector("#RejestrIntrodukcji1_JXPGrid_RI_cell_0_0").click()

            # click print button
            try:
                d_elem = WebDriverWait(browser, delay).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#RejestrIntrodukcji1_SmallRibbonButton13_Text")))
                print("ready to print")
            except TimeoutException:
                print("too long")
            browser.find_element_by_css_selector("#RejestrIntrodukcji1_SmallRibbonButton13_Text").click()

            # find printing window
            r = None
            icon_to_click = "Printing window"
            while r is None:
                r = pyautogui.locateOnScreen(bmp_print, grayscale=True)
                time.sleep(0.3)
            print((icon_to_click) + ' Visible')
            buttoon_print = pyautogui.locateCenterOnScreen(bmp_print, grayscale=True)
            print(buttoon_print)
            pyautogui.moveTo(buttoon_print)
            pyautogui.doubleClick()
            pyautogui.typewrite("2")
            pyautogui.press('enter')
            pyautogui.moveRel(50, 0)
            time.sleep(1)

            time_now = datetime.datetime.now()
            time_var = time.strftime('%H%M%S')
            log_file.write(time_var + ' ' + str(single_number) + '\n')

        except:
            print("Cannot find accession number")
        browser.switch_to.default_content()

    log_file.close()

    auto_increment = 's'


# ------------------------------------ #

# open Firefox
browser.get('https://mycompany.com')

# log in
browser.find_element_by_id("TextBox1").send_keys(login)
browser.find_element_by_id("TextBox2").send_keys(password)
browser.find_element_by_id("TextBox2").send_keys(Keys.ENTER)

# Do you need autoincrement?
auto_increment = input('Do you need autoincrement?\nInput yes or no\n').lower()

# remember response
if auto_increment == "yes":
    response = 'yes'
elif auto_increment == "no":
    response = "no"

# make sure you get right answer "yes", "no" or "q"
while auto_increment or len(auto_increment) == 0:
    if auto_increment == "yes" or auto_increment == "no":
        scanning()
        if auto_increment != 'q':
            print_barcodes(accession_numbers)
    elif auto_increment == 's':
        continue_scan = input('Continue?\nInput \"yes\" lub \"Q\", to quit.\n').lower()
        if continue_scan == 'yes':
            # show input
            auto_increment = response
        elif continue_scan == 'q':
            auto_increment = 'q'
        else:
            continue_scan = input('\nInput \"yes\", to continue or \"Q\", to quit\n').lower()
    elif auto_increment == 'q':
        break
    else:
        auto_increment = input('Please input correct value.\nDo you need autoincrement?\ninput yes or no\n').lower()
