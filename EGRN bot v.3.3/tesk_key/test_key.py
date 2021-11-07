import os, re, sys

AuthKey = None
##------------------------------------------------------------##
def EGRNauth():
    global AuthKey
    encoding = None
    try:
        with open("auth.txt", "rb") as f:
##        with open("auth-8+.txt", "rb") as f:
##        with open("auth-16-be.txt", "rb") as f:
##        with open("auth-16-le.txt", "rb") as f:
            BOM = f.read(2)
        if BOM == b'\xfe\xff': encoding = "UTF-16 BE"
        elif BOM == b'\xff\xfe':  encoding = "UTF-16 LE"
        elif BOM == b'\xef\xbb':  encoding = "UTF-8 со спецификацией"
        
        if encoding is None:
            with open("auth.txt", "r") as f:
                AuthKey = f.read(36)
            if re.search("\w{8}-\w{4}-\w{4}-\w{4}-\w{12}", AuthKey) != None: AuthStatus = "GotAuthKey"
            else: return "NoAuthKey"
        else:
            print(f"Файл сохранён с кодировкой \"{encoding}\".")
            return "ErrAuthKeyEncoding"

        return "GotAuthKey"

    except:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(f"При выполнении возникло исключение {exc_type}\n"
            f"\tописание:\t{exc_obj}\n"
            f"\tмодуль:\t\t{fname}\tстрока:\t{exc_tb.tb_lineno}")
        return "ErrAuthKeyExc"

Status = EGRNauth()
if Status == "GotAuthKey": print("Файл в порядке, ключ правильный.")
if Status == "NoAuthKey": print("Нет ключа доступа. Дальнейшая работа невозможна.")
if Status == "ErrAuthKeyEncoding": print("Проблема с кодировкой файла. Дальнейшая работа невозможна.")
if Status == "ErrAuthKeyExc": print("Проблема с файлом. Дальнейшая работа невозможна.")

