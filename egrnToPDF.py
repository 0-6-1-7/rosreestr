# https://habr.com/ru/post/523368/
# https://habr.com/ru/post/479978/

import logging
import os
import platform
import shutil
import time
import zipfile
from random import randint
from time import sleep

from selenium.webdriver import Chrome
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

from config import CaptchaReloadingCounterMax, SaveCaptcha, isHeadless
from tools.recognize import recognize

DEBUG = True

if DEBUG:
    logging.basicConfig(
        format=u'%(filename)-18s [LINE:%(lineno)-4s] %(levelname)-8s [%(asctime)s] %(message)s',
        level=logging.INFO,

    )
else:
    logging.basicConfig(
        format=u'%(filename)-18s [LINE:%(lineno)-4s] %(levelname)-8s [%(asctime)s] %(message)s',
        level=logging.WARNING
        # filename='debug.log'
    )


def wait(a, b=0):
    if b == 0:
        b = a
    sleep(randint(a, b))


class Parser:
    def __init__(self, name=''):
        # self.name = name
        # self.config = self.getConfig()
        # self.cities = self.config['cities'].split(', ')

        # random.shuffle(self.cities)
        self.isHeadless = isHeadless
        self.sourceDir = os.chdir('Responses')
        self.sourcePath = os.path.abspath(os.getcwd())
        self.folders = self.getZips()
        self.options = self.optionsDriver()
        self.driver = self.initDriver()
        self.getURL('https://rosreestr.gov.ru/wps/portal/cc_vizualisation')
        self.treatment()
        #

    def RecognizeCaptcha(self, CaptchaImage):
        try:
            img_base64 = self.driver.execute_script(
                """var ele = arguments[0];var cnv = document.createElement('canvas');cnv.width = 180; cnv.height = 50;cnv.getContext('2d').drawImage(ele, 0, 0);return cnv.toDataURL('image/png').substring(22);""",
                CaptchaImage)
            captcha = recognize(img_base64, SaveCaptcha)
        except:
            logging.warning("Ошибка при обработке капчи. Скорее всего, это ошибка сайта.")
            captcha = ""
        if captcha == "44444":
            logging.warning("Ошибка при обработке капчи (44444).")
            captcha = ""
        return captcha

    def getCaptcha(self):
        ##    print("Обработка капчи")
        OK = False
        global CaptchaReloadingCounter

        while True:
            time.sleep(1)
            try:
                CaptchaImage = self.driver.find_element_by_xpath('//*[@id="captchaImage2"]')
                # CaptchaImage = self.driver.find_element_by_xpath(".//img[@style='width: 180px; height: 50px;']")
                # CaptchaImage = self.driver.find_element_by_id('captchaImage2')
                # CaptchaImage = vapp.find_element_by_id()
            except:
                logging.warning("На странице не найдена капча")
                break
            captcha = self.RecognizeCaptcha(CaptchaImage)
            if captcha != "":
                return captcha
            CaptchaReloadingCounter = CaptchaReloadingCounter + 1
            print(f"     получение другой капчи (попытка № {CaptchaReloadingCounter} из {CaptchaReloadingCounterMax})")
            CaptchaReloadButton = self.driver.find_element_by_xpath(
                ".//span[contains(@class,'v-button-caption') and contains(text(),'Другую картинку')]")
            CaptchaReloadButton.click()
            if CaptchaReloadingCounter > CaptchaReloadingCounterMax:
                print("Капча не получена после нескольких попыток. Вероятно, ошибка сайта.")
                break
            time.sleep(10)
        return ("00000")

    def getZips(self):
        folders = []

        for filename in os.listdir(self.sourceDir):
            dir = filename.split('.zip')[0]
            try:
                os.mkdir(dir)
                logging.info(f'Make dir for response: {dir}')
            except Exception as ex:
                pass

            if filename.endswith('.zip'):
                shutil.move(filename, f'{dir}/{filename}')
                logging.info(f'Extractall: {dir}/{filename} to {dir}')
                zipfile.ZipFile(f'{dir}/{filename}', 'r').extractall(f'{dir}')
                # os.remove(filename)
                os.remove(f'{dir}/{filename}')

            folders.append(dir)

        logging.info(folders)
        return folders

    def treatment(self):
        for folder in self.folders:
            logging.info(f'Get sig, zip, xml from {folder}')
            files = os.listdir(folder)
            logging.info(files)
            # sig_file = f'{self.sourcePath}/{folder}/{files[0]}'
            # zip_file = f'{self.sourcePath}/{folder}/{files[1]}'

            # logging.info(f'Extractall zip_file')
            # zipfile.ZipFile(zip_file, 'r').extractall(folder)

            # logging.info('Get xml')
            for f in os.listdir(folder):
                if f.endswith('.zip'):
                    zip_file = f'{self.sourcePath}/{folder}/{f}'
                    zipfile.ZipFile(zip_file, 'r').extractall(folder)
                if f.endswith('.sig'):
                    sig_file = f'{self.sourcePath}/{folder}/{f}'

            for f in os.listdir(folder):
                if f.endswith('.xml'):
                    xml_file = f'{self.sourcePath}/{folder}/{f}'

            # logging.info(f'Send sig_file')
            act = self.driver.find_element_by_id('sig_file')
            act.send_keys(sig_file)

            # logging.info(f'Send xml_file')
            act = self.driver.find_element_by_id('xml_file')
            act.send_keys(xml_file)

            captcha = self.getCaptcha()
            logging.info(f'Обработаем капчу: {captcha}')

            self.driver.find_element_by_css_selector('input.brdg1111').send_keys(captcha)

            self.driver.find_element_by_xpath("//button[text()='Проверить ']").click()

            act = self.driver.find_element_by_link_text('Показать в человекочитаемом формате')
            act.click()

            wait(1)
            handles = self.driver.window_handles
            size = len(handles)
            # print(size)
            # for x in range(size):
            self.driver.switch_to.window(handles[1])

            # print('-' * 30)
            # print(f'!!! {self.driver.title}')

            element = self.driver.find_element_by_class_name('noprint')
            self.driver.execute_script("""
            var element = arguments[0];
            element.parentNode.removeChild(element);
            """, element)

            with open(f'{folder}/{folder}.html', 'w') as f:
                f.write(self.driver.page_source)

            self.driver.close()

            self.driver.switch_to.window(handles[0])

    def optionsDriver(self):
        options = Options()
        options.add_argument('window-size=1024,768')

        if self.isHeadless:
            options.add_argument('-headless')

        # options.add_argument('log-level=3')

        # prefs = {"profile.managed_default_content_settings.images": 2}
        # options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])

        options.add_argument('--disable-gpu')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-plugins')
        options.add_argument('--disable-application-cache')
        options.add_argument('--disable-infobars')
        # options.add_argument('--no-sandbox')
        # options.add_argument('-incognito')

        userAgent = 'Mozilla/5.0 (Windows NT 10.0; rv:78.0) Gecko/20100101 Firefox/78.0'
        options.add_argument(f'user-agent={userAgent}')
        options.add_argument("--disable-blink-features=AutomationControlled")

        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument('--disable-logging')

        return options

    def initDriver(self):
        logging.info('Init Driver')

        try:
            if platform.system() == 'Linux':
                return Chrome('drivers/chromedriverUnix', options=self.options)
                logging.info(f'Init driver: chromedriverUnix')
            if platform.system() == 'Darwin':
                return Chrome('drivers/chromedriverDarwin', options=self.options)
                logging.info(f'Init driver: chromedriverDarwin')
            else:
                return Chrome('drivers/chromedriver', options=self.options)
                logging.info(f'Init driver: chromedriver')

        except Exception as ex:
            logging.warning(f'Use ChromeDriverManager')
            # logging.warning(f'{ex}')
            return Chrome(ChromeDriverManager().install(), options=self.options)

    def getURL(self, url):
        self.driver.get(url)
        logging.info(f'Get page - {url}')
        logging.info(f'current_url - {self.driver.current_url}')

        if self.driver.current_url != url:
            timeOut = randint(5, 12)

            logging.warning(f'Failed to load')
            logging.warning(f"Pass url: {url}")
            logging.warning(f"Current_url: {self.driver.current_url}")
            logging.warning(f'Wait {timeOut}sec')

            sleep(timeOut)
            self.getURL(url)

        return self.driver


if __name__ == '__main__':
    logging.info('----------------')
    logging.info('Parser is start')
    parser = Parser()
    # logging.info(f"Total projects: {projects}")
    try:
        parser.driver.quit()
        logging.info(f'Driver is quit\n')
    except Exception as ex:
        logging.warning(f'Driver already closed')
        logging.warning(f'{ex}')
