from datetime import datetime

from os import path

from re import match, search, MULTILINE, IGNORECASE

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import InvalidSessionIdException, NoSuchElementException

from sys import exc_info

from time import monotonic, sleep

# ------------------------------------------------------------ #
def get_header():
    return convert("header")

# ------------------------------------------------------------ #
def convert(content):
    try:
        rs = [
            (1, r"(?:Статус объекта\n)(.*)(?:\n)", "Статус объекта"),
            (2, r"(?:Дата обновления информации:\n)(\d\d\.\d\d\.\d{4})(?:\n)", "Дата обновления информации"),
            (3, r"(?:Вид объекта недвижимости\n)(.*)(?:\n)", "Вид объекта недвижимости"),
            (4, r"(?:Площадь,.*\n)(.*)(?:\n)", "Площадь"),
            (5, r"(?:Площадь, )(.*)(?:\n)", "Ед. изм."),
            (6, r"(?:Адрес[а]* \(местоположение\)\n)(.*)(?:\n)", "Адрес"),
            (7, r"(?:(?:Назначение)\n|(?:Вид разрешенного использования)\n)(.*)(?:\n)", "Назначение или вид разрешенного использования"),
            (8, r"(?:Вид, номер и дата государственной регистрации права\n)((?:.*\n)*)", "Дата последенего перехода права")
            ]

    ### дополнительные сведения
    ### порядковый номер, регулярное выражение для разбора данных, заголовок
        rs.append((9, r"(?:Форма собственности\n)(.*)(?:\n)", "Форма собственности"))
        rs.append((10, r"(?:Этаж\n)(.*)(?:\n)", "Этаж"))
        rs.append((11, r"(?:Кадастровая стоимость \(руб\)\n)(.*)(?:\n)", "Кадастровая стоимость (руб)"))
        rs.append((12, r"(?:Вид, номер и дата государственной регистрации права\n)((?:.*\n)*?)"\
            "(?:(?:Ограничение прав и обременение объекта недвижимости)|(?:ВЕРНУТЬСЯ К РЕЗУЛЬТАТАМ ПОИСКА))",
            "Вид, номер и дата государственной регистрации права")) ##только права
        rs.append((13, r"(?:Ограничение прав и обременение объекта недвижимости\n)((?:.*\n)*?)"\
            "(?:ВЕРНУТЬСЯ К РЕЗУЛЬТАТАМ ПОИСКА)",
            "Ограничение прав и обременение объекта недвижимости")) ##только ограничения
        rs.append((14, r"(?:Вид, номер и дата государственной регистрации права\n)((?:.*\n)*)"\
            "(?:(?:Ограничение прав и обременение объекта недвижимости)|(?:ВЕРНУТЬСЯ К РЕЗУЛЬТАТАМ ПОИСКА))",
            "Сведения о правах и ограничениях (обременениях)")) ##права и ограничения
    ###
        
        t = ""
        if content == "header":
            for r in rs:
                t = t + ("\t", "")[len(t) == 0] + r[2]
            return t
        for r in rs:
            sr = search(r[1], content, flags=MULTILINE | IGNORECASE)
            if sr: s = sr.group(1)
            else: s = "Нет данных"

            if r[0] == 1 and search(r"Актуально", s) is None: t = s + "\t" * 8; break

            if r[0] == 4: # площадь - может быть записана как через ".", так и через ","
                if s != "Нет данных":
                    s = s.replace(".", ",")

            if r[0] == 8:
                if s != "Нет данных":
                    textlines = s.splitlines()
                    dl = []
                    for n in range(0, len(textlines)):
                        if match(r"(^Ограничение)|(^ВЕРНУТЬСЯ)", textlines[n], flags=IGNORECASE) is not None: break
                        dstr = search(r"от (\d{2}\.\d{2}\.\d{4})", textlines[n])
                        if dstr is not None: dl.append(datetime.strptime(dstr.group(1), "%d.%m.%Y"))
                    if len(dl) == 0:
                        s = "дата не указана"
                    else:
                        dl.sort(reverse = True)
                        s = dl[0].strftime("%d.%m.%Y")

            t = t + ("\t", "")[len(t) == 0] + s

    except:
        exc_type, exc_obj, exc_tb = exc_info()
        fname = path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(f"При выполнении возникло исключение {exc_type}\n"
            f"\tописание:\t{exc_obj}\n"
            f"\tмодуль:\t\t{fname}\tстрока:\t{exc_tb.tb_lineno}")

        t = ""
    return t

# ------------------------------------------------------------ #
def wait_timeout(timeout_start=0, timeout=5, minimal_timeout=2, msg=""):
    print("Таймаут ", end='')
    t0 = monotonic()
    if t0 + minimal_timeout > timeout_start + timeout:
        print(f"{minimal_timeout} сек от текущего времени ", end='')
    else:
        print(f"{timeout} сек от предыдущего запроса ", end='')
    
    while monotonic() <= max(t0 + minimal_timeout, timeout_start + timeout):
        if msg: print('>', end='')
        sleep(1)

    print(f" {msg}")

# ------------------------------------------------------------ #
def check_notifications(RR, silent=False):
    notifications = RR.find_element(By.CSS_SELECTOR, "div.notifications")
    if not notifications: return False
    
    buttons = notifications.find_elements(By.CSS_SELECTOR, "button.rros-ui-lib-button.rros-ui-lib-button--link")
    if buttons:
        if not silent: print(f"сообщений об ошибках: {len(buttons)}")
        return notifications
    else:
        return False

# ------------------------------------------------------------ #
def clear_notifications(RR, silent=False):
    notifications = check_notifications(RR, silent=True)
    if not notifications: return False

    messages = notifications.find_elements(By.CSS_SELECTOR, "div.rros-ui-lib-error-message")
    last_notification = messages[-1].text
    print(f"последнее сообщение = {last_notification}");

    buttons = notifications.find_elements(By.CSS_SELECTOR, "button.rros-ui-lib-button.rros-ui-lib-button--link")
    for button in buttons:
        print(button)
        button.click()
        
    return last_notification

# ------------------------------------------------------------ #
## ожидание элемента или сообщения об ошибке
def alternative_wait(RR, timeout, by_locator, locator, text_1, text_2):
    t0 = monotonic()
    while monotonic() < t0 + timeout:
        sleep(1)
        if check_notifications(RR):
            last_notification = clear_notifications(RR)
            return "NOTIF_alternative_wait"

        ## элемент, который на самом деле нужен, а не вот это вот всё 
        try:
            element = RR.find_element(by_locator, locator)
            if (text_1 in element.text) or (text_2 in element.text):
                return True
        except:
            print("Исключение в alternative_wait")

    return False
    
# ------------------------------------------------------------ #
def wait_spinner_on_off(RR):
    wait_3 = WebDriverWait(RR, 3)
    wait_60 = WebDriverWait(RR, 60)

    t0 = monotonic()
    ## ждём появления блока спиннера
    try:
        spinner = wait_3.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div.rros-ui-lib-spinner__wrapper")))
        print(f"\tспиннер появился за {round(monotonic() - t0, 6)} сек.")
    except:
        spinner = False
        print("Исключение _ON_ в wait_spinner_on_off")

    t0 = monotonic()
    ## ждём пропадания блока спиннера
    nospinner = False
    if spinner:
        try:
            nospinner = wait_60.until(
                EC.staleness_of(spinner))
            print(f"\tспиннер пропал за {round(monotonic() - t0, 1)} сек.")
        except:
            print("Исключение _OFF_ в wait_spinner_on_off")

    # True - если спиннер нормально появился и пропал
    return nospinner

# ------------------------------------------------------------ #
def get_info_kn(RR, kn, last_query_time, timeout):
    ## last_query_time - точка отсчёта для таймаута 5 секунд
    
    wait_3 = WebDriverWait(RR, 3)
    wait_60 = WebDriverWait(RR, 60)


    try: #wrap for any return
        ## проверям заголовок страницы
        try:
            sipono = wait_3.until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, "div.rros-ui-lib-title")))
            page_title = sipono.get_attribute("innerText")
            if page_title != "Справочная информация по объектам недвижимости в режиме online":
                return "FATAL_ERROR_PAGE_UNKNOWN"
        except:
            return "FATAL_ERROR_PAGE_UNKNOWN"

        ## ждём появления кнопки поиска (как индиактор полной загрузки страницы)
        search_button = wait_60.until(
            EC.visibility_of_element_located(
                (By.ID, "realestateobjects-search")))
        if search_button is None:
            return "FATAL_ERROR_PAGE_NOT_FULLY_LOADED"

        print("getinfo - проверка наличия сообщения об ошибках - перед началом поиска")
        last_notification = clear_notifications(RR)
        if last_notification:
##            print("sleep(10)"); sleep(10)
            return "NOTIF_" + last_notification
        
        ## ввод кадастрового номера в поле
        kn_input = RR.find_element(
            By.ID, "query")
        kn_input.click()

        if kn_input.get_attribute("value") != "":
            kn_input.send_keys(
                Keys.BACKSPACE * len(kn_input.get_attribute("value")))

        kn_input.send_keys(kn)

        ## ждём активизации кнопки поиска
        try:
            search_button_active = wait_3.until(
                EC.element_to_be_clickable(search_button))
        except:
            search_button_active = False
        if not search_button_active:
            return "ERROR_TIMEOUT" # кнопка поиска неактивна - видимо, какая-то ошибка

        t0 = monotonic()
        print("\tперед таймаутом:")
        print(f"\t\tlast_query_time = {last_query_time}")
        print(f"\t\tmonotonic() = {t0}")
        wait_timeout(
            timeout_start=last_query_time,
            minimal_timeout=0,
            timeout=timeout,
            msg="таймаут перед нажатием кнопки поиска")
        print(f"\t\tожидание = {round(monotonic() - t0, 2)}")
        print(f"\t\tс предыдущего запроса = {round(monotonic() - last_query_time, 2)}")
        
## запускаем поиск
        # 5 попыток выполнить поиск
        retry_count = 0
        while True:
            search_button.click()
            print("Нажата нопка поиска")
            t0 = monotonic()

            spinner_ok = wait_spinner_on_off(RR)
            wait_timeout
            ## проверяем наличие сообщения об ошибке
            print("getinfo - проверка наличия сообщения об ошибках - после окончания поиска")
            last_notification = clear_notifications(RR)
            if last_notification:
                return "NOTIF_" + last_notification

            if not spinner_ok: # спиннер крутился слишком долго или другая проблема со спиннером
               return "ERROR_SEARCH_TIMEOUT"

## поиск закончился, обрабатываем результат
            print(f"\tПоиск завершился за {round(monotonic() - t0, 1)} сек.")

            result = alternative_wait(RR, 60,
                                      By.CSS_SELECTOR, "div.row.realestateobjects-wrapper__results",
                                      kn, "По заданным критериям поиска объекты не найдены")

            if not result:
                return "ERROR_SEARCH_FAILED"

            if type(result) == type("NOTIF") and result.startswith("NOTIF"):
                retry_count += 1
                print(f"Попытка запустить поиск {retry_count}")
                sleep(retry_count)
                if retry_count  > 4:
                    return result
            else:
                break

        results = RR.find_element(
                By.CSS_SELECTOR, "div.row.realestateobjects-wrapper__results")

        if "По заданным критериям поиска объекты не найдены" in results.text:
            return "Объект по кадастровому номеру не найден"

        t0 = monotonic()
        print("\tперед таймаутом:")
        print(f"\t\tlast_query_time = {last_query_time}")
        print(f"\t\tmonotonic() = {t0}")
        wait_timeout(
            timeout_start=last_query_time,
            timeout=timeout,
            minimal_timeout=3,
            msg="перед открытием карточки")
        print(f"\t\tожидание = {round(monotonic() - t0, 2)}")
        print(f"\t\tс предыдущего запроса = {round(monotonic() - last_query_time, 2)}")
        
        data_rows = RR.find_elements(
            By.CSS_SELECTOR, "div.rros-ui-lib-table__row")

        # 5 попыток открыть карточку
        retry_count = 0
        while True:
            data_rows[0].click()
            print("Открывается карточка объекта")
            t0 = monotonic()

            ## ожидание карточки объекта
            object_info_window = alternative_wait(RR, 60,
                                                  By.CSS_SELECTOR, "div.rros-ui-lib-modal__window_full_screen",
                                                  kn, "По заданным критериям поиска объекты не найдены")
            
            if not object_info_window:
                return "ERROR_KN_FOUND_NOT_SHOWN"

            if type(object_info_window) == type("NOTIF") and object_info_window.startswith("NOTIF"):
                retry_count += 1
                print(f"Попытка открыть карточку {retry_count}")
                sleep(retry_count)
                if retry_count  > 4:
                    return object_info_window
            else:
                break

        object_info_window = RR.find_element(By.CSS_SELECTOR, "div.rros-ui-lib-modal__window_full_screen")
        
        ## зафикисруем время для отсчёта нового таймаута, оно будет возвращено в основную программу
        last_query_time = monotonic()

        ## обработка карточки объекта
        object_info = convert(object_info_window.get_attribute("innerText"))

        ## ожидание сслыки "вернуться к результатам поиска"
        reverse_link = wait_3.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button.realestate-object-modal__btn.rros-ui-lib-button--reverse")))
        reverse_link.click()
        print(f"\tКарточка объекта обработана за {round(monotonic() - t0, 1)} сек.")
        
        return {"object_info":object_info, "last_query_time":last_query_time}

    except:
        print("произошло исключение в get_info_kn")
        exc_type, exc_obj, exc_tb = exc_info()
        fname = path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(f"При выполнении возникло исключение {exc_type}\n"
            f"\tописание:\t{exc_obj}\n"
            f"\tмодуль:\t\t{fname}\tстрока:\t{exc_tb.tb_lineno}")
        return  "NOTIF_произошло исключение в get_info_kn"

# ------------------------------------------------------------ #
def get_info_addr(RR, data, D):
    try: #wrap for any return
        wait = WebDriverWait(RR, 5)
        try: sipono = wait.until(EC.visibility_of_element_located((By.ID, "sipono-selector")))
        except: return "FATAL_ERROR_PAGE_UNKNOWN"

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
                exc_type, exc_obj, exc_tb = exc_info()
                fname = path.split(exc_tb.tb_frame.f_code.co_filename)[1]
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
        exc_type, exc_obj, exc_tb = exc_info()
        fname = path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(f"При выполнении возникло исключение {exc_type}\n"
            f"\tописание:\t{exc_obj}\n"
            f"\tмодуль:\t\t{fname}\tстрока:\t{exc_tb.tb_lineno}")
        return  "NOTIF_произошла какая-то хрень в get_info_addr"
