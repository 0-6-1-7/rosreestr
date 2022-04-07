import re

def clear_address_for_search(addr):
    a = addr
    
    a = re.sub(r"\bВолгоград\b.*р-н \S*\,", "Волгоград", a, flags=re.IGNORECASE)
    a = re.sub(r"\bкв\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bд\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bвлд\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bдвлд\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bим\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bимени\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bкорп\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bк\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bком\b.*", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bобл\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bобласть\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bр-н\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bрайон\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bроссия\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bгенерала\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bмаршала\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bакадемика\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bмкр\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bмикрорайон\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"\bтер\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"ё", "е", a, flags=re.IGNORECASE)
    a = re.sub(r"-й\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"-ый\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"-я\b", " ", a, flags=re.IGNORECASE)
    a = re.sub(r"-ая\b", " ", a, flags=re.IGNORECASE)

## специальные случаи
    a = re.sub(r"волгоград\b.*\bпос.*?\b", " ", a, flags=re.IGNORECASE)
## 400020, Волгоградская обл, Волгоград г, р-н Кировский, тер Поселок им Саши Чекалина, д. 83, кв. 7
## 400032, Волгоградская обл, Волгоград г, р-н Кировский, ул ПОС ВЕСЕЛАЯ БАЛКА, д. 45, кв. 24



    return a
