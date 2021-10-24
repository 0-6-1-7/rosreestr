import re
from datetime import datetime

def Convert(content):
    rs = [
        (1, r"(?:Дата обновления информации\:\n)(\d\d\.\d\d\.\d{4})(?:\n)"),
        (2, r"(?:Вид объекта недвижимости\n)(.*)(?:\n)"),
        (3, r"(?:Статус объекта\n)(.*)(?:\n)"),
        (4, r"(?:Площадь,.*\n)(.*)(?:\n)"),
        (5, r"(?:Площадь, )(.*)(?:\n)"),
        (6, r"(?:Адрес[а]* \(местоположение\)\n)(.*)(?:\n)"),
        (7, r"(?:(?:Назначение)\n|(?:Вид разрешенного использования)\n)(.*)(?:\n)"),

        (8, r"(?:Вид, номер и дата государственной регистрации права\n)((?:.*\n)*)")
        ]
    t = ""
    for r in rs:
        try: s = re.search(r[1], content, flags = re.MULTILINE | re.IGNORECASE).group(1)
        except: s = "Ошибка при получении информации"
        
        if r[0] == 3 and re.search(r"Актуально", s) == None: t = t + "\t" + s + "\t" * 5; break

        if r[0] == 8:
            if s == "Ошибка при получении информации":
                s = "Сведения о переходах прав отсутствуют"
            else:
                s = re.sub(r"\n(от \d{2}\.\d{2}\.\d{4})", " \g<1>", s, flags = re.MULTILINE | re.IGNORECASE)
                textlines = s.splitlines()
                n = 0
                dl = []
                while n < len(textlines):
                    if re.match(r"(^Ограничение)|(^ВЕРНУТЬСЯ)", textlines[n], flags = re.IGNORECASE) == None:
                        dstr = re.search(r"\d{2}\.\d{2}\.\d{4}", textlines[n + 1])
                        if dstr == None: s = "Невеозможно получить дату перехода права"; break
                        d = dstr.group(0)
                        dl.append(datetime.strptime(d, "%d.%m.%Y"))
                        n = n + 2
                    else: break
                dl.sort(reverse = True)
                s = dl[0].strftime("%d.%m.%Y")

        t = t + ("\t", "")[len(t) == 0] + s

    return t
