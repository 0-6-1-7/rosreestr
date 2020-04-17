from recognize import recognize
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl, time
import re

EGRN = None
status = 0 # состояние сайта
  
##------------------------------------------------------------##
def EGRNinit():
    global EGRN
    EGRN = webdriver.Chrome()
    time.sleep(1)
    EGRN.set_page_load_timeout(10)
    try:
        EGRN.get("https://rosreestr.ru/wps/portal/p/cc_present/ir_egrn")
        EGRN.implicitly_wait(1)
    except:
        print("Сайт Росреестра не работает")
##------------------------------------------------------------##
def RecognizeCaptcha(c):
    global EGRN
    img_base64 = EGRN.execute_script("""var ele = arguments[0];var cnv = document.createElement('canvas');cnv.width = 180; cnv.height = 50;cnv.getContext('2d').drawImage(ele, 0, 0);return cnv.toDataURL('image/png').substring(22);""", c)
    captcha = recognize(img_base64)
    return captcha
  
##------------------------------------------------------------##
def GetInfo():
    global EGRN
    vapp = EGRN.find_element_by_css_selector("div.v-app")
    KN = re.search(r"\b\d{2}:\d{2}:\d{1,7}:\d{1,}\b", EGRN.find_element_by_css_selector("div.header3").get_attribute("innerText"))[0]
    CaptchaImage = vapp.find_element_by_xpath("//img[@style='width: 180px; height: 50px;']")
    captcha = RecognizeCaptcha(CaptchaImage)
    Radio = vapp.find_elements_by_xpath("//input[@type='radio']")[1]
    CaptchaField = vapp.find_elements_by_xpath("//input[@type='text']")[1]
    SendButton = vapp.find_element_by_xpath("//span[contains(@class,'v-button-caption') and contains(text(),'Отправить запрос')]")
    EGRN.execute_script("arguments[0].scrollIntoView();", Radio)
    Radio.click()
    CaptchaField.click()
    time.sleep(1)
    CaptchaField.send_keys(captcha)
    time.sleep(1)
    SendButton.click()
    PopUp = []
    while len(PopUp) == 0:
##        print("     no popup")
        time.sleep(2)
        PopUp = vapp.find_elements_by_xpath("//div[@class='v-window']")
        time.sleep(2)
        if PopUp[0].get_attribute("innerText") == "Ошибка запроса\nПревышен интервал между запросами":
            PopUp = []
            print("     popup = timeout")
##    time.sleep(2)
    OkButton = PopUp[0].find_element_by_xpath("//span[contains(@class,'v-button-caption') and contains(text(),'Продолжить работу')]")
    NZ = re.search(r"\d{2}\-\d{9}", PopUp[0].get_attribute("innerText"))[0]
    OkButton.click()
    print(f"Кадастровый номер {KN} запрос {NZ}")
##    time.sleep(1)
##    CloseButton = vapp.find_element_by_xpath("//span[contains(@class,'v-button-caption') and contains(text(),'Закрыть')]")
##    CloseButton.click()
    print("Пауза между запросами 5 минут")
    for t in range(0, 5):
        time.sleep(60)
        print(f"     прошло {t + 1} минут")
##    print("Ура!")
    
    
##------------------------------------------------------------##
def GetAll():
    global EGRN
    AllKN = []
    AllON = []
    DataTable = EGRN.find_element_by_css_selector("table.v-table-table")
    AllON = DataTable.find_elements_by_css_selector("tr > td:first-child")
    for ON in AllON:
        AllKN.append(ON.get_attribute("innerText"))
    for KN in AllKN:
        DataTable = EGRN.find_element_by_css_selector("table.v-table-table")
        AllON = []
        AllON = DataTable.find_elements_by_css_selector("tr > td:first-child")
        for ON in AllON:
            if KN == ON.get_attribute("innerText"):
                EGRN.execute_script("arguments[0].scrollIntoView();", ON)
                ON.click()
                time.sleep(2)
                GetInfo()
                break
    print("Всё готово!")

##------------------------------------------------------------##
EGRNinit()
