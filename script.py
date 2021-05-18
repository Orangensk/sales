import pandas
from collections import defaultdict, Counter
from openpyxl import load_workbook, Workbook

# читаем эксель файл
logs = pandas.read_excel('logs.xlsx', sheet_name='log')

# Определяем константы
ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUFV'

# Инициализируем словарь для хранения данных о продажах товаров в разбивке по полам 'м' и 'ж'
dict_of_sales = {
  'м': [],
  'ж': []
}

# преобразуем в словарь
logs_dict = logs.to_dict(orient='records')

# создаём словарь браузеров
browser_dict = defaultdict(int)

# создаём словарь Браузеры - по месяцам
browsers_by_month = defaultdict(int)

# создаём словарь товары - по месяцам
items_by_month = defaultdict(int)

# Инициализируем плоский список для браузеров. Одна строка файла - один элемент списка
list_of_browsers = []

# Обход данных из Excel
for log_record in logs_dict:

    # Получаем месяц
    month = log_record['Дата посещения'].month

    # Купленные товары делим на элементы
    for item in log_record['Купленные товары'].split(','):
        
        # Добавляем товар в словарь dict_of_sales
        dict_of_sales[log_record['Пол']].append(item.strip())

        items_by_month[item.strip()+ '=' + str(month)] += 1

    # Добавляем браузер в словарь browsers_by_month
    browsers_by_month[log_record['Браузер']+'='+str(month)] += 1

    # Добавляем название браузера в плоский список
    list_of_browsers.append(log_record['Браузер'])

# Открываем файл отчета report.xlsx
wb = load_workbook(filename='report.xlsx')
sheet = wb['Лист1']

# Выводим в отчет 7 самых популярных браузеров
# Начальная строка - 5
row = 5

# Получем 7 попялярных браузеров и обходим каждый из них
most_common = Counter(list_of_browsers).most_common(7)
for item in most_common:
   
  # Начальная колонка - 2
  col = 2

  # Получаем название браузера
  browser = item[0]

  # Записываем название браузера в первюю колонку - колонку A
  sheet['A'+str(row)] = browser

  # Запускаем цикл от 1 до 12 - по количеству месяцев в году
  for i in range(1, 13):

    # Получаем количество посещений текущего браузера за текущий месяц из словаря browsers_by_month
    # Если такого ключа в словаре нет, то передаем 0
    quantity = browsers_by_month.get(browser+'='+str(i), 0)

    # Записываем количество в ячейку
    sheet[ALPHABET[col]+str(row)] = quantity
    col += 1

  row += 1

# Выполняем для товаров тоже самое, что делали для браузеров
row = 19
counter = Counter(dict_of_sales['м']+dict_of_sales['ж'])
for item in counter.most_common(7):
  col = 2
  title = item[0]
  quantiti = item[1]
  sheet[ALPHABET[0]+str(row)] = title
  for i in range(1, 13):
    sheet[ALPHABET[col]+str(row)] = items_by_month.get(title+'='+str(i), 0)
    col += 1
  row += 1

# Получаем самые популярные товары среди мужчин
counter_male = Counter(dict_of_sales['м'])
# Прямой порядок - самый популярный товар
counter_male_most_common = counter_male.most_common(1)[0]
# Обратный порядок - самый невостребованный товар
counter_male_most_common_reverse = counter_male.most_common()[:-(len(counter_male)+1):-1][0]

# Получаем самые популярные товары среди мужчин
counter_female = Counter(dict_of_sales['ж'])
counter_female_most_common = counter_female.most_common(1)[0]
counter_female_most_common_reverse = counter_female.most_common()[:-(len(counter_female)+1):-1][0]

sheet['B31'] = f'{counter_male_most_common[0]} ({counter_male_most_common[1]})'
sheet['B32'] = f'{counter_female_most_common[0]} ({counter_female_most_common[1]})'
sheet['B33'] = f'{counter_male_most_common_reverse[0]} ({counter_male_most_common_reverse[1]})'
sheet['B34'] = f'{counter_female_most_common_reverse[0]} ({counter_female_most_common_reverse[1]})'

wb.save('report_output.xlsx')
