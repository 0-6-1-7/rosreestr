### Обновление 11.09.2024

Добавлено:

- кнопка для быстрой вставки содежимого бучера обмена в поле поиска; если в буфере обмена содержится кадастровый номер с пробелами, они удалются (' 34:35:000000 :1234 ' > '34:35:000000:1234')

![Кнопка для быстрой вставки ](https://github.com/0-6-1-7/rosreestr/blob/master/Chrome%20extension%20v.3.1/screenshots/6.png)

### Обновление 09.09.2024

Добавлено:

- кнопки для быстрого копирования кадастрового номера, сведений о правах и всей информации;

![Кнопки для быстрого копироания ](https://github.com/0-6-1-7/rosreestr/blob/master/Chrome%20extension%20v.3.1/screenshots/5.png)

### Обновление 02.09.2024

Добавлено:

- более яркая подсветка в списке результатов;
- если на экране отображается карточка объекта, то при переключении режима работы автоматически на лету копируется кадастровый номер или вся информации в зависимости от режима, ранее нужно было выйти из карточки и открыть её повторно.

### Обновление 01.09.2024

Переписано расширение для Chrome.

Добавлено:

- переключение режима работы: выключено, копирование кадастрового номера, копирование всей информации - переключается циклически Выкл > 34 > i > Выкл и т.д.
  
![Режим работы](https://github.com/0-6-1-7/rosreestr/blob/master/Chrome%20extension%20v.3.1/screenshots/4.png)
- подсветка статуса объекта: актуально + есть сведения о правах (зелёный), актуально + нет сведений о правах (синий), погашено (красный)
- подсветка кадастрового номера при копировании: скопирован (зелёный), не скопирован, т.к. окно не в фокусе (красный) - нужно вернуться в окно
![Список найденных объектов](https://github.com/0-6-1-7/rosreestr/blob/master/Chrome%20extension%20v.3.1/screenshots/3.png)

### Обновление 23.08.2024

Добавлено расширение для Chrome.

Скрывает лишние элементы, делает отображение информации более компактным, автоматически закрывает сообщения об ошибках.

При просмотре сведений об объекте если объект имеет статус "Актуально" и есть раздел со сведениями о правах, автоматически копирует кадастровый номер в буфер обмена и подсвечивает его зелёным фоном.

Внешний вид страниц с работающим расширением и без него (в обоих случаях установлен масштаб страницы 90%)

![Список найденных объектов](https://github.com/0-6-1-7/rosreestr/blob/master/Chrome%20extension%20v.3.1/screenshots/1.png)

![Сведения об объекте](https://github.com/0-6-1-7/rosreestr/blob/master/Chrome%20extension%20v.3.1/screenshots/2.png)



### Обновление 09.08.2024

Переработана процедура авторизации, описание в файkе esia.ini

Исправлно нексолько опчеаток.


### Обновление 02.08.2024

**LK SIpON** - сбор информации по кадастровым номерам

**LK EGRN** - заказ выписок из ФГИС ЕГРН
    
    Основные скрипты:
    - LK EGRN about - сведения об объекте,
    - LK EGRN transfer - сведения о переходе прав на объект.
    
    Скрипт загрузки результатов:
    - начать с https://script.google.com/, 
    - создать проект,
    - пройти через несколько этапов авторизации,
    - запустить.
    Скрипт читает письма от росреестра, собирает из них ссылки на загрузку результатов и присылает все в одном письме. 
    Обрабатываются только непрочитанные письма. После обработки письма помечаются звёздочкой, устанавливается отметка о прочтении.

    Все старые скрипты перемещены в ветку "Устарело"


========================

**Устарело**

========================

**EGRN LK bot** - бот для заказа выписок из ФГИС ЕГРН - новая версия сервиса в ЛК (см. новость от 13.12.22)

    Тестовая версия. 
    Сделана автоматическая авторизация через ЕСИА - описание в файле esia.ini
    Запрос выписок работает, никакие ошибки не обрабатываются. 
    Загрузка результатов будет позже.

**EGRN bot** - бот для запроса выписок из ФГИС ЕГРН - старая версия сервиса (см. новость от 13.12.22)

**EGRN addr** - скрипт для поиска кадастровых номеров по адресам во ФГИС ЕГРН

~~SIpON v4 LK - скрипт для сбора информации по адресам или кадастровым номерам в личном кабинете Росреестра~~
В связи с ноябрьскими изменениями в ЛК перестал работать SIpON v4 LK. 

**SIpON v5 LK** - скрипт для сбора информации ~~по адресам или~~ кадастровым номерам в личном кабинете Росреестра
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

