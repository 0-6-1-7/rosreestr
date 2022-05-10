PILOK = False
SeleniumOK = False
OpenpyxlOK = False
ChromeDriverOK = False
ChromeOK = False

print("\n\n\nПроверка необходимых компонентов\n\n\n")

try:
    import PIL
    PILOK = True
    print("+ Библиотека PIL установлена и работает")
except: print("- Ошибка: библиотека PIL не установлена или не работает")

try:
    import selenium
    SeleniumOK = True
    print("+ Библиотека selenium установлена и работает")
except: print("- Ошибка: библиотека selenium не установлена или не работает")

try:
    import openpyxl
    OpenpyxlOK = True
    print("+ Библиотека openpyxl установлена и работает")
except: print("- Ошибка: библиотека openpyxl не установлена или не работает")



import os, subprocess
try:
    ChromeDriverVersion = subprocess.run(["ChromeDriver", "-v"], stdout = subprocess.PIPE).stdout.decode('utf-8').split(" ")[1]
    ChromeDriverOK = True
    print(f"+ ChromeDriver установлен и доступен")
except:
    print("- ChromeDriver не установлен или недоступен в PATH:")
    print("\t" + os.environ.get("PATH").replace(";", "\n\t"))

if SeleniumOK:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    try:
        Chrome = webdriver.Chrome(options = chrome_options)
        Chrome.get("chrome://version") # можно запросить любую страницу, если не возникнет исключение, значит версии Crhrome и ChromeDriver совместимы
        ChromeOK = True
        print("+ Версии Crhrome и ChromeDriver совместимы")
        Chrome.close(); Chrome.quit()
    except:    
        print("- Ошибка: версии Crhrome и ChromeDriver несовместимы")

if PILOK and SeleniumOK and OpenpyxlOK and ChromeDriverOK and ChromeOK: print("\n\n\nВсё хорошо. Можно закрыть окно.")
else: print("\n\n\nОбнаружены проблемы. Дальнейшая работа, скорее всего, невозможна.")

import time
time.sleep(1000)
