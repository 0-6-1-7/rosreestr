import re
from datetime import datetime

def Convert(content):
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
        try: s = re.search(r[1], content, flags = re.MULTILINE | re.IGNORECASE).group(1)
        except: s = "Не указано для этого ОН или ошибка"
        
        if r[0] == 1 and re.search(r"Актуально", s) == None: t = s + "\t" * 6; break

        if r[0] == 8:
            if s != "Не указано для этого ОН или ошибка":
                textlines = s.splitlines()
                dl = []
                for n in range(0, len(textlines)):
                    if re.match(r"(^Ограничение)|(^ВЕРНУТЬСЯ)", textlines[n], flags = re.IGNORECASE) != None: break
                    dstr = re.search(r"от (\d{2}\.\d{2}\.\d{4})", textlines[n])
                    if dstr != None: dl.append(datetime.strptime(dstr.group(1), "%d.%m.%Y"))
                if len(dl) == 0:
                    s = "дата не указана"
                else:
                    dl.sort(reverse = True)
                    s = dl[0].strftime("%d.%m.%Y")

        t = t + ("\t", "")[len(t) == 0] + s

    return t
