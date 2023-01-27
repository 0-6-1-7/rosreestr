**EGRN LK bot** - бот для заказа выписок из ФГИС ЕГРН - новая версия сервиса в ЛК (см. новость от 13.12.22)

    Тестовая версия. 
    Сделана автоматическая авторизация через ЕСИА - описание в файле esia.ini
    Запрос выписок работает, никакие ошибки не обрабатываются. 
    Загрузка резульатов будет позже.

**EGRN bot** - бот для запроса выписок из ФГИС ЕГРН - старая версия сервиса (см. новость от 13.12.22)

**EGRN addr** - скрипт для поиска кадастровых номеров по адресам во ФГИС ЕГРН

~~SIpON v4 LK - скрипт для сбора информации по адресам или кадастровым номерам в личном кабинее Росреестра~~
В связи с ноябрьскими изменениями в ЛК перестал работать SIpON v4 LK. 

**SIpON v5 LK** - скрипт для сбора информации ~~по адресам или~~ кадастровым номерам в личном кабинее Росреестра
Тестовая версия


### Новости по работе с ФГИС ЕГРН по состоянию на 27.01.23

Примерно с 15 января сервис стал работать крайне ненадёжно. Немного поправить ситуацию может следующее:
1. В тексте скриптов EGRN bot нужно заменить адреса "rosreestr.ru" на "rosreestr.gov.ru"
2. В тексте скриптов EGRN bot нужно изменить значения констант (строки 20, 21)

    **defaultBatchMax = 1**

    **defaultBatchSize = 1**


### Новости по работе с ФГИС ЕГРН по состоянию на 13.12.22

1. С 09.12.22 Росреестр перешёл на использование Российских цифровых сертификатов (т.н. «цифровой суверенитет»). В связи с этим сайт перестал открываться во всех браузерах, кроме Российских Яндекс.Браузер и Атом, в которые интегрированы корневые сертификаты Минцифры.
Для продолжения работы с сайтом в Google Chrome нужно вручную добавить корневой сертификат Минцифры в хранилище Windows.
Как это сделать - в инструкции [Работа с ФГИС ЕГРН («старый» способ) после 9 декабря 2022.docx](https://github.com/0-6-1-7/rosreestr/blob/master/%D0%A0%D0%B0%D0%B1%D0%BE%D1%82%D0%B0%20%D1%81%20%D0%A4%D0%93%D0%98%D0%A1%20%D0%95%D0%93%D0%A0%D0%9D%20(%C2%AB%D1%81%D1%82%D0%B0%D1%80%D1%8B%D0%B9%C2%BB%20%D1%81%D0%BF%D0%BE%D1%81%D0%BE%D0%B1)%20%D0%BF%D0%BE%D1%81%D0%BB%D0%B5%209%20%D0%B4%D0%B5%D0%BA%D0%B0%D0%B1%D1%80%D1%8F%202022.docx)

2. Сообщение с официальной страницы из Личного кабинета Росреестра:

>Сервис «Запрос посредством доступа к ФГИС ЕГРН» модернизирован
>
>С 21 ноября 2022 года введен в действие модернизированный сервис «Запрос посредством доступа к ФГИС ЕГРН».
>Пополнение баланса и подача заявок будет осуществляться посредством обновленной версии сервиса, размещенной в Личном кабинете Официального сайта Росреестра. Предыдущая версия сервиса доступна до 21 ноября 2023 года в открытой части Официального сайта Росреестра (https://rosreestr.gov.ru/wps/portal/p/cc_present/ir_egrn).
>
>Внимание!
>
>Баллы по пакетам сервиса «Запрос посредством доступа к ФГИС ЕГРН», оплаченные до его модернизации, могут использоваться только в предыдущей версии сервиса. Проверить >баланс Вы можете в Личном кабинете в разделе «Мой баланс» – вкладка «Архив».
>
>Баллы, оплаченные после 21 ноября 2022 года, могут использоваться только в обновленной версии сервиса, размещенной в Личном кабинете.
>Подробное описание особенностей функционирования обновленного сервиса «Запрос посредством доступа к ФГИС ЕГРН» доступно в Руководстве пользователя Личного кабинета.

В связи с п.2 идёт работа по созданию нового бота. Ориентировный срок выхода тестовой версии - до нового года.
