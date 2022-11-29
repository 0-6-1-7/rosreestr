import os, re, sys, time

from datetime import datetime

from time import monotonic, sleep

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import InvalidSessionIdException, NoSuchElementException

# ------------------------------------------------------------ #
def get_header():
    return convert("header")

# ------------------------------------------------------------ #
def convert(content):
    rs = [
        (1, "(?:Статус объекта\n)(.*)(?:\n)", "Статус объекта"),
        (2, "(?:Дата обновления информации\:\n)(\d\d\.\d\d\.\d{4})(?:\n)", "Дата обновления информации"),
        (3, "(?:Вид объекта недвижимости\n)(.*)(?:\n)", "Вид объекта недвижимости"),
        (4, "(?:Площадь,.*\n)(.*)(?:\n)", "Площадь"),
        (5, "(?:Площадь, )(.*)(?:\n)", "Ед. изм."),
        (6, "(?:Адрес[а]* \(местоположение\)\n)(.*)(?:\n)", "Адрес"),
        (7, "(?:(?:Назначение)\n|(?:Вид разрешенного использования)\n)(.*)(?:\n)", "Назначение или вид разрешенного использования"),
        (8, "(?:Вид, номер и дата государственной регистрации права\n)((?:.*\n)*)", "Дата последенего перехода права")
        ]

### дополнительные сведения
### порядковый номер, регулярное выражение для разбора данных, заголовок
    rs.append((9, "(?:Форма собственности\n)(.*)(?:\n)", "Форма собственности"))
    rs.append((10, "(?:Этаж\n)(.*)(?:\n)", "Этаж"))
    rs.append((11, "(?:Кадастровая стоимость \(руб\)\n)(.*)(?:\n)", "Кадастровая стоимость (руб)"))
    rs.append((12, "(?:Вид, номер и дата государственной регистрации права\n)((?:.*\n)*?)"\
        "(?:(?:Ограничение прав и обременение объекта недвижимости)|(?:ВЕРНУТЬСЯ К РЕЗУЛЬТАТАМ ПОИСКА))",
        "Вид, номер и дата государственной регистрации права")) ##только права
    rs.append((13, "(?:Ограничение прав и обременение объекта недвижимости\n)((?:.*\n)*?)"\
        "(?:ВЕРНУТЬСЯ К РЕЗУЛЬТАТАМ ПОИСКА)",
        "Ограничение прав и обременение объекта недвижимости")) ##только ограничения
    rs.append((14, "(?:Вид, номер и дата государственной регистрации права\n)((?:.*\n)*)"\
        "(?:(?:Ограничение прав и обременение объекта недвижимости)|(?:ВЕРНУТЬСЯ К РЕЗУЛЬТАТАМ ПОИСКА))",
        "Сведения о правах и ограничениях (обременениях)")) ##права и ограничения
###
    
    t = ""
    if content == "header":
        for r in rs:
            t = t + ("\t", "")[len(t) == 0] + r[2]
        return t
    for r in rs:
        try: s = re.search(r[1], content, flags=re.MULTILINE | re.IGNORECASE).group(1)
        except: s = "Не указано для этого ОН или ошибка"

        if r[0] == 1 and re.search(r"Актуально", s) is None: t = s + "\t" * 8; break

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
def get_info_kn(RR, kn):
    try: #wrap for any return
        wait = WebDriverWait(RR, 5)
        print(1)
        try: sipono = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.rros-ui-lib-title")))
        except: return "FATAL_ERROR_UNKNOWN_PAGE"
        print(2)
        page_title = sipono.get_attribute("innerText")
        if page_title != "Справочная информация по объектам недвижимости в режиме online":
            return "FATAL_ERROR_UNKNOWN_PAGE"
        print(3)
        kn_input = RR.find_element(By.ID, "query")
        
        n = 1
        kn_input.click()
        if kn_input.get_attribute("value") != "":
            kn_input.send_keys(Keys.BACKSPACE * len(kn_input.get_attribute("value")))

        try:
            prev_results = RR.find_element(By.CSS_SELECTOR, "div.realestateobjects-wrapper__results")
        except:
            prev_results = None

        kn_input.send_keys(kn)
        search_button = RR.find_element(By.ID, "realestateobjects-search")
        search_button.click()

        t0 = monotonic()
        while True:
            results = RR.find_element(By.CSS_SELECTOR, "div.realestateobjects-wrapper__results")
            results_text = results.get_attribute("innerText")
            if kn in results_text: break
            if "По заданным критериям поиска объекты не найдены" in results_text and results != prev_results:
                return (kn, None)
            sleep(1)
            if monotonic() - t0 > 60: return "ERROR_SEARCH_FAILED"


        data_rows = RR.find_elements(By.CSS_SELECTOR, "div.rros-ui-lib-table__row")
        data_rows[0].click()


        wait = WebDriverWait(RR, 60)
        try:
            object_info_window = wait.until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, "div.rros-ui-lib-modal__window_full_screen")))
        except:
            return "ERROR_KN_FOUND_NOT_SHOWN"

        object_info = convert(object_info_window.get_attribute("innerText"))
        reverse_link = wait.until(
            EC.visibility_of_element_located(
                (By.CSS_SELECTOR, "button.realestate-object-modal__btn.rros-ui-lib-button--reverse")))
        reverse_link.click()
        

        print("get_info_kn отработало нормально")
        return object_info

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
        sleep(1)
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
