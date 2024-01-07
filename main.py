import pandas as pd
import requests
from tkinter import Tk, Label, Entry, Button
from urllib.parse import unquote

def download_xlsx():
    url = entry.get()
    # Проверяем, является ли ссылка Google Таблицей
    if "docs.google.com/spreadsheets/" not in url:
        print("Некорректная ссылка на Google Таблицу")
        return

    try:
        # Отправляем GET-запрос по указанной ссылке
        response = requests.get(url)

        # Проверяем, успешно ли получен ответ
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Ошибка при получении Google Таблицы: {e}")
        return

    # Имя файла - последняя часть ссылки
    table_id = url.split("/")[-2]
    print(table_id)

    download_url = f'https://docs.google.com/spreadsheets/d/{table_id}/export?format=xlsx'

    # Отправляем GET-запрос и сохраняем файл
    response = requests.get(download_url)

    filename = unquote(response.headers.get('Content-Disposition').split('filename*=UTF-8\'\'')[1]).replace(';', '_')

    # Скачивание таблицы и сохранение ее в файл
    response = requests.get(download_url)
    with open(filename, 'wb') as file:
        file.write(response.content)

    print(f"Google Таблица успешно скачана и сохранена в файле: {table_id}")

# Создание пользовательского интерфейса
root = Tk()
root.title('Скачать Google Таблицу')

label = Label(root, text='Введите ссылку на Google Таблицу:')
label.pack()

entry = Entry(root, width=50)
entry.pack()

button = Button(root, text='Скачать', command=download_xlsx)
button.pack()

root.mainloop()


# Считываем файлы в pandas DataFrame
df_xlsx = pd.read_excel('ИКТК-11 (ТКП).xlsx', header=1)
df_csv = pd.read_csv('stat_231218_18-22.csv', header=0)

# Сравниваем столбцы
df_xlsx.columns = [col.strip() for col in df_xlsx.columns]
df_csv.columns = [col.strip() for col in df_csv.columns]
print(df_xlsx.columns)
print(df_csv.columns)

# Сравнение столбцов
common_cols = [col for col in df_xlsx.columns if col in df_csv.columns]
df_common_cols = df_csv[common_cols]
print(common_cols)

# Загрузка данных из xlsx
try:
    xlsx_data = pd.read_excel('ИКТК-11 (ТКП).xlsx')
    xlsx_sheets = pd.ExcelFile("ИКТК-11 (ТКП).xlsx").sheet_names
    # Обновление данных, если вкладка уже существует

    if 'SOTSBI' in xlsx_sheets:
        with pd.ExcelWriter("ИКТК-11 (ТКП).xlsx", mode="a", if_sheet_exists="replace") as writer:
            df_common_cols.to_excel(writer, sheet_name='SOTSBI', index=False, columns=df_csv.columns)

    # Добавление новой вкладки, если вкладка не существует
    else:
        with pd.ExcelWriter('ИКТК-11 (ТКП).xlsx', engine='openpyxl', mode='a') as writer:
            df_common_cols.to_excel(writer, sheet_name='SOTSBI', index=False)
except FileNotFoundError:
    # Создание нового файла xlsx, если файл не существует
    with pd.ExcelWriter('ИКТК-11 (ТКП).xlsx') as writer:
        df_common_cols.to_excel(writer, sheet_name='SOTSBI', index=False)
print('Таблица успешно добавлена в файл xlsx')