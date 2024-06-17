# Yandex_Market - Процесс по сбору информации из "Яндекс Маркета"

### Общая информация ###

+ Расположение процесса Yandex_Market : C:\Users\%USERNAME%\Documents\Primo\Yandex_Market
+ Расположение макроса: C:\Users\%USERNAME%\Documents\Primo\Yandex_Market\Scripts\FindMinimumCost.bas
+ Расположение Config: C:\Users\%USERNAME%\Documents\Primo\Yandex_Market\Data\Config.xls
+ Расположение итогового файла : C:\Users\%USERNAME%\Documents\Primo\Yandex_Market\Data\Итоговый файл.xlsx"

### Краткое описание робота ###

**Процесс состоит из следующих частей**
1. Робот открывает Яндекс Маркет.
2. Робот осуществляет поиск по уже заданному товару
3. Робот собирает информацию об указанном товаре на первой странице
4. Робот записывает таблицу с товарами в Excel файл, и находит строчку с минимальной стоимостью.

**Подробнее**
1. Робот открывает Google Chrome, и переходит на страницу Яндекс Маркета. Предусмотрены проверки на открытие сайта и на наличие окна с авторизацией(которое нужно закрыть). Всего есть три попытки открыть Яндекс маркет
2. Робот в поиске указывает товар, который ранее указан в Config файле. Предусмотрена проверка на то, что мы действительно нашли нужный товар, а также на то, что данный товар существует. Всего есть три попытки найти нужный товар.
3. Робот отображает все товары в виде таблицы, где в каждом ряду находится четыре товара. Затем в цикле мы проходимся по всем таким группам товаров. Предусмотрен случай, когда цена на товар указана с оплатой картой Пэй или её отсутствием.
4. В End Process, мы записываем DataTable с информацией по товарам в Excel файл. Затем, с помощью скрипта FindMinimumCost мы находим строчку с минимальной стоимостью товара и подкрашиваем её. 
5. После сохранения результирующего файла, мы отправляем файл по почте с использованием протокола SMTP. Для этого, необходимо разкомментировать модуль ReadPasswords где происходит чтение паролей, добавить креды в диспетчер учетных данных,
и разкомментировать модуль по отправке письма SendResultFile
### Как это работает. ###

1. **INITIALIZATION**
  + ./Framework/*InitiAllSettings* - Загрузка данных конфигурации из Config.xlsx файл и из активов
  + ./Framework/*CreateTransactions* - Создание транзакций для процесса в переменную TransactionDataDT
  + ./Framework/*ReadPasswords* - Чтение паролей из диспетчера паролей windows, ассеты которых записанны в словаре Assets

2. **GET TRANSACTION DATA**
  + ./Framework/*GetTransactionData* - Получение транзакции TransactionItem (строки) из TransactionData. 

3. **PROCESS TRANSACTION**
  + *Process* - Отслеживание процесса и вызов других рабочих процессов, связанных с автоматизируемым процессом. 
  + ./Framework/*SetTransactionStatus* - Обновляет статус обработанной транзакции : Успех, бизнес-исключение или системное исключение
  + ./Processes/Browser/ChechinkProductInfo.ltw - Сбор информации по выыбранному товару и запись его в ProductsInformationDT
  + ./Processes/Browser/FindRequiredProduct.ltw - Поиск товара, который указан в Config файле
  + ./Processes/Browser/OpenYandexMarket.ltw - Открытие сайта Яндекс Маркета через Google Chrome


4. **END PROCESS**
  + ./Framework/*KillAllProcesses* - Закрытие Google Chrome и Excel
  + ./Framework/*SendResultFile* - Отправка результирующего файла по протоколу SMTP
  + ./Processes/Excel/SaveResultFile.ltw - Запись ProductsInformationDT в Excel файл, и обработка его макросом FindMinimumCost (поиск наименьшей стоимости)