import os, re, sys, time

from datetime import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import InvalidSessionIdException, NoSuchElementException

# ------------------------------------------------------------ #
def convert(content):
    rs = [
        (1, r"(?:Статус объекта\n)(.*)(?:\n)"),
        (2, r"(?:Дата обновления информации\:\n)(\d\d\.\d\d\.\d{4})(?:\n)"),
        (3, r"(?:Вид объекта недвижимости\n)(.*)(?:\n)"),
        (4, r"(?:Площадь,.*\n)(.*)(?:\n)"),
        (5, r"(?:Площадь, )(.*)(?:\n)"),
        (6, r"(?:Адрес[а]* \(местоположение\)\n)(.*)(?:\n)"),
        (7, r"(?:(?:Назначение)\n|(?:Вид разрешенного использования)\n)(.*)(?:\n)"),

        (8, r"(?:Вид, номер и дата государственной регистрации права\n)((?:.*\n)*)")
        ]
    t = ""
    for r in rs:
        try: s = re.search(r[1], content, flags=re.MULTILINE | re.IGNORECASE).group(1)
        except: s = "Не указано для этого ОН или ошибка"

        if r[0] == 1 and re.search(r"Актуально", s) is None: t = s + "\t" * 7; break

        if r[0] == 8:
            if s != "Не указано для этого ОН или ошибка":
                textlines = s.splitlines()
                dl = []
                for n in range(0, len(textlines)):
                    if re.match(r"(^Ограничение)|(^ВЕРНУТЬСЯ)", textlines[n], flags=re.IGNORECASE) is not None: break
                    dstr = re.search(r"от (\d{2}\.\d{2}\.\d{4})", textlines[n])
                    if dstr is not None: dl.append(datetime.strptime(dstr.group(1), "%d.%m.%Y"))
                if len(dl) == 0:
                    s = "дата не указана"
                else:
                    dl.sort(reverse = True)
                    s = dl[0].strftime("%d.%m.%Y")

        t = t + ("\t", "")[len(t) == 0] + s

    return t

# ------------------------------------------------------------ #
def get_info_kn(RR, D):
    try: #wrap for any return
        wait = WebDriverWait(RR, 5)
        try: sipono = wait.until(EC.visibility_of_element_located((By.ID, "sipono-selector")))
        except: return "FATAL_ERROR_UNKNOWN_PAGE"

        kn_input = sipono.find_element(By.CSS_SELECTOR, "input")
        
        n = 1
        while kn_input.get_attribute("value") != "":
            try:
                kn_input.send_keys(Keys.ENTER)
                kn_inputClear = wait.until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "div.rros-ui-lib-dropdown__clear-indicator")))
                kn_inputClear.click()
            except:
                kn_input.send_keys(Keys.ESCAPE)
                kn_input.send_keys(Keys.END)
                kn_input.send_keys(Keys.BACKSPACE * len(kn_input.get_attribute("value")))
            n = n + 1
            if n > 5: return "ERROR_CAN_NOT_CLEAR_INPUT_FIELD"
        KN10 = ";".join(key for key, value in D.items()) + ";"
        results = None
        kn_input.send_keys(KN10)
        time.sleep(1)

        wait = WebDriverWait(RR, 60)
        try:
            results = wait.until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, "div.realestateobjects-wrapper__results")))
        except:
            return "ERROR_KN_FOUND_NOT_SHOWN"

        wait = WebDriverWait(RR, 5)
        pagination = results.find_element(
            By.CSS_SELECTOR, "div.rros-ui-lib-table-pagination__count")
        pages_count = int(pagination.get_attribute("innerText").split(" ")[3])
        if pages_count > 1:
            nextpage_button = results.find_element(
            By.CSS_SELECTOR, "button.rros-ui-lib-table-pagination__btn--next")

        prev_resultrows = None
        for page in range(1, pages_count + 1):
            if page > 1: nextpage_button.click()
            if prev_resultrows is not None:
                while True:
                    resultrows = results.find_element(
                        By.CSS_SELECTOR, "div.rros-ui-lib-table__rows")
                    if prev_resultrows.get_attribute("innerText") != resultrows.get_attribute("innerText"):
                        prev_resultrows = resultrows; break

            resultrows = results.find_elements(
                By.CSS_SELECTOR, "div.rros-ui-lib-table__row")
            if len(resultrows) > 0:
                for row in resultrows:
                    kn = row.find_element(
                        By.CSS_SELECTOR, "div > div > a").get_attribute("innerText")
                    row.click()
                    try:
                        reverselink = wait.until(
                            EC.visibility_of_element_located(
                                (By.CSS_SELECTOR, "button.realestate-object-modal__btn.rros-ui-lib-button--reverse")))
                        objectinfo = wait.until(
                            EC.visibility_of_element_located(
                                (By.CSS_SELECTOR, "div.realestate-object-modal")))
                    except:
                        return "ERROR_OBJECT_NOT_SHOWN"
                    info = convert(objectinfo.get_attribute("innerText"))
                    RR.execute_script("arguments[0].click();", reverselink)
                    D[kn] = info

        print("get_info_kn отработало нормально")
        return "OK"

    except:
        print("произошло исключение в get_info_kn")
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(f"При выполнении возникло исключение {exc_type}\n"
            f"\tописание:\t{exc_obj}\n"
            f"\tмодуль:\t\t{fname}\tстрока:\t{exc_tb.tb_lineno}")
        return  "NOTIF$произошла какая-то хрень в get_info_kn"

# ------------------------------------------------------------ #
def get_info_addr(RR, data, D):
    try: #wrap for any return
        wait = WebDriverWait(RR, 5)
        try: sipono = wait.until(EC.visibility_of_element_located((By.ID, "sipono-selector")))
        except: return "FATAL_ERROR_UNKNOWN_PAGE"

        data_input = sipono.find_element(By.CSS_SELECTOR, "input")
        n = 1
        while data_input.get_attribute("value") != "":
            try:
                data_input.send_keys(Keys.ENTER)
                data_input_clear = wait.until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "div.rros-ui-lib-dropdown__clear-indicator")))
                data_input_clear.click()
            except:
                data_input.send_keys(Keys.ESCAPE)
                data_input.send_keys(Keys.END)
                data_input.send_keys(Keys.BACKSPACE * len(data_input.get_attribute("value")))
            n = n + 1
            if n > 5: return "ERROR_CAN_NOT_CLEAR_INPUT_FIELD"

        data_input.send_keys(data)
        time.sleep(1)
        wait = WebDriverWait(RR, 30) ##10

        while True:
            loo = wait.until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, "div.rros-ui-lib-dropdown__menu")))
            loaded = wait.until(
                EC.invisibility_of_element_located(
                    (By.CSS_SELECTOR, "rros-ui-lib-dropdown__loading-indicator")))
            objects = loo.get_attribute("innerText").split("\n")
            if objects[0] != "Поиск объектов недвижимости...": break

        if objects[0] == "Информация не найдена":
            data_input.send_keys(Keys.END)
            data_input.send_keys(Keys.BACKSPACE * len(data_input.get_attribute("value")))
            data_input.send_keys(Keys.ESCAPE)
        else:
            data_input.send_keys(Keys.ENTER)
            try:
                data_input_clear = wait.until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "div.rros-ui-lib-dropdown__clear-indicator")))
                data_input_clear.click()
            except:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                return "error"

        if objects[0] == "Информация не найдена":
            sta = f"Информация не найдена\t"
        else:
            info = ""
            objects_used = min(30, len(objects))
            print(f"Найдено: {len(objects)}, будет обработано: {objects_used}")
            for i in range(0, objects_used, 2):
                info = info + ("" if len(info) == 0 else "\n") + objects[i + 1] + "\t" + objects[i]
            info = "\n".join(set(info.split("\n")))
            
            for d in set(info.split("\n")):
                D[d.split("\t")[0]] = ""
            sta = get_info_kn(RR, D)
        return sta

    except:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(f"При выполнении возникло исключение {exc_type}\n"
            f"\tописание:\t{exc_obj}\n"
            f"\tмодуль:\t\t{fname}\tстрока:\t{exc_tb.tb_lineno}")
        return  "NOTIF$произошла какая-то хрень в get_info_addr"
