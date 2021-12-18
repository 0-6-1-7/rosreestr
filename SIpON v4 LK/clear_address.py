import re

def clear_address_for_search(addr):
    a = addr
    a = re.sub(r"([\ \.,]кв[\.\ ])", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"([\ \.,]д[\.\ ])", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"([\ \.,]им[\.\ ])", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"([\ \.,]корп[\.\ ])", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"([\ \.,]к[\.\ ])", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"([\ \.,]ком[\.\ ])", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"ё", "е", a, flags=re.IGNORECASE)

    return a
