import pandas as pd
import requests
from tkinter import Tk, Label, Entry, Button, filedialog
from urllib.parse import unquote
import pickle
import os

filename = ''


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

    table_id = url.split("/")[-2]

    download_url = f'https://docs.google.com/spreadsheets/d/{table_id}/export?format=xlsx'

    # Отправляем GET-запрос и сохраняем файл
    response = requests.get(download_url)

    global filename
    filename = unquote(response.headers.get('Content-Disposition').split('filename*=UTF-8\'\'')[1]).replace(';', '_')

    # Скачивание таблицы и сохранение ее в файл
    response = requests.get(download_url)
    with open(filename, 'wb') as file:
        file.write(response.content)

    print(f"Google Таблица успешно скачана и сохранена в файле: {filename}")


file_path = ''


def browse_file():
    global file_path
    file_path = filedialog.askopenfilename()

    # Сохранение пути к файлу в файле pickle
    with open('file_path.pickle', 'wb') as fp:
        pickle.dump(file_path, fp)

    # Обновление текста Label с выбранным путем и названием файла
    file_label.config(text=f'Выбран файл: {os.path.basename(file_path)}')

    pass


# Создание пользовательского интерфейса
root = Tk()
root.title('Скачать Google Таблицу')

label = Label(root, text='Введите ссылку на Google Таблицу:')
label.pack()

entry = Entry(root, width=50)
entry.pack()

# Создание Label для отображения выбранного пути и названия файла
file_label = Label(root, text='')
file_label.pack()

button_browse = Button(root, text='Обзор', command=browse_file)
button_browse.pack()

button_download = Button(root, text='Скачать', command=download_xlsx)
button_download.pack()

# Попытка загрузить путь к файлу из pickle, если он существует
try:
    with open('file_path.pickle', 'rb') as f:
        file_path = pickle.load(f)
        file_label.config(text=f'Выбран файл: {os.path.basename(file_path)}')
except FileNotFoundError:
    file_path = None

root.mainloop()

# Считываем файлы в pandas DataFrame
df_xlsx = pd.read_excel(filename, header=1)
df_csv = pd.read_csv(file_path, header=0)

# Сравниваем столбцы
df_xlsx.columns = [col.strip() for col in df_xlsx.columns]
df_csv.columns = [col.strip() for col in df_csv.columns]

common_cols = [col for col in df_xlsx.columns if col in df_csv.columns]
df_common_cols = df_csv[common_cols]

# Загрузка данных из csv в xlsx
try:
    xlsx_data = pd.read_excel(filename)
    xlsx_sheets = pd.ExcelFile(filename).sheet_names
    # Обновление данных, если лист уже существует

    if 'SOTSBI' in xlsx_sheets:
        with pd.ExcelWriter(filename, mode="a", if_sheet_exists="replace") as writer:
            df_common_cols.to_excel(writer, sheet_name='SOTSBI', index=False, columns=df_csv.columns)

    # Добавление нового листа, если лист не существует
    else:
        with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
            df_common_cols.to_excel(writer, sheet_name='SOTSBI', index=False)
except FileNotFoundError:
    # Создание нового файла xlsx, если файл не существует
    with pd.ExcelWriter(filename) as writer:
        df_common_cols.to_excel(writer, sheet_name='SOTSBI', index=False)
print(f'Таблица успешно добавлена в {filename}') в файл xlsx')
