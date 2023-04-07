import datetime, json, re

def Convert(content):
  rs = [
    (1, r"Кадастровый номер:\t(\d\d:\d\d:\d+:\d+.*)\n"),
    (2, r"Дата обновления информации:\t(.*?)\n"),
    (3, r"^(.*?)\n"),
    (4, r"Статус объекта:\t(.*?)\n"),
    (5, r"Площадь.*:\t(.*?)\n"),
    (6, r"(?:(?:Единицы измерения площади)|(?:Единица измерения \(код\))):\t(.*?)\n"),    
    (7, r"Адрес[а]* \(местоположение\):\t(.*?)\n"),
    (8, r"\(ОКС\) Тип:\t(.*?)\n"),
    (9, r"Форма собственности:\t(.*?)\n"),
    (10, r"\nПрава и ограничения\n((.|\n)*?)\n(Вернуться к результатам поиска)")
    ]
  t = ""
  for r in rs:
    try:
      s = re.search(r[1], content, re.MULTILINE).group(1)
    except:
      s = "ошибка при получении информации"
      if r[0] == 4:
        s = "Данные не актуализированы"
      elif r[0] == 8:
        s = "Данные не актуализированы"
    finally:
      if len(s) == 0:
        if r[0] == 9:
          s = "Форма собств. не указана"
        elif r[0] == 10:
          s = "Данные отсутствуют"
      else:
        s = re.sub("(\t| |\n)+\n", "\n", s)
      if r[0] == 10:
        if re.search(r"собственность|управление", s, flags = re.MULTILINE | re.IGNORECASE) != None:
          regex = r"(\d\d\.\d\d\.\d\d\d\d).*?(собственность|управление)"
          matches = re.finditer(regex, s, flags = re.MULTILINE | re.IGNORECASE)
          dl = []
          for m in matches:
            d = m.group(1)
            f = "%d.%m.%Y"
            dobj = datetime.datetime.strptime(d, f)
            dl.append(dobj)  
          dl.sort(reverse = True)
          s = dl[0].strftime("%d.%m.%Y")
        else:
          s = "Данные отсутствуют"
    s = s.strip()
    t = t + s + "\t"
  return t.strip(" \t")
