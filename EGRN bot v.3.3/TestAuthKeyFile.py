import re

def TestAuthKeyFile():
    try:
        f = open("auth.txt", "r"); AuthKey = f.readline(); f.close()
        print("Файл имеет правильное имя и доустпен для чтения, прочитана первая строка, всё ОК.")
    except:
        print("Ошибка: нет нужного файла или к нему нет доступа.")
        return
    try:
        k = re.search("\w{8}-\w{4}-\w{4}-\w{4}-\w{12}", AuthKey)
        if k.start() == 0:
            print ("Ключ найден с первой позиции, всё ОК.")
        else:
            print (f"Ошибка, ключ найден с позиции {k.start()}")
    except:
        print("Ошибка, ключ не найден в певрвой строке.")
## ---------------------------------------------------
        
TestAuthKeyFile()
