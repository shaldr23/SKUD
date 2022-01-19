# %%
"""
Автор: Шапошников А.В.
Новая версия программы. Исходные файлы - pdf в виде таблицы
Старая версия перенесена в /old_versions
"""

import os
import subprocess
import sys
import pandas as pd
import numpy as np
import re
from datetime import datetime


# ------------------ Functions -------------------------------------------

def writelog(text):
    """
    Function to securely write information into log file (logfilename)
    """
    with open(logfile, 'a') as f:
        f.write(text + '\n')


def make_df_from_pdf(input_file: str) -> pd.DataFrame:
    """
    1. Используется конвертер pdftotext.exe, вывод (текст) хранится в виде строки
    2. Строка парсится и переводится в DataFrame
    3. Даты и время форматируются.
    4. В столбце "ДАТА" вставляются пропущенные значения из другого столба
    """
    pdftotext_process = subprocess.run(f'pdftotext.exe -table -enc UTF-8 "{input_file}" "-"',
                                       shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                       encoding=sys.stdout.encoding)
    name_pattern = r'(?:([А-Я]\S+)\s+)' * 3
    datetime_pattern = r'(\S+)\s*' * 4
    pattern = name_pattern + r'(?:(?:Мужской)|(?:Женский))\s+' + datetime_pattern
    data = []
    for line in pdftotext_process.stdout.split('\n'):
        found = re.search(pattern, line)
        if found:
            groups = found.groups()
            name = ' '.join(groups[:3])
            data.append((name, *groups[3:]))
    frame = pd.DataFrame(data, columns=('ФИО', 'ДАТА', 'ПРИХОД', 'ДАТА_2', 'УХОД'))
    frame[['ДАТА', 'ДАТА_2']] = frame[['ДАТА', 'ДАТА_2']].apply(pd.to_datetime, format='%d.%m.%Y', errors='coerce')
    frame[['ПРИХОД', 'УХОД']] = frame[['ПРИХОД', 'УХОД']].apply(pd.to_datetime, format='%H:%M:%S', errors='coerce')
    frame['ДАТА'] = frame['ДАТА'].fillna(frame['ДАТА_2'])
    frame.drop(columns='ДАТА_2', inplace=True)
    return frame


def get_staff_from_file(file: str) -> pd.Series:
    """
    Получить список сотрудников из Табеля.
    Ищутся файлы табеля .xlsx в папке data/info и парсятся
    """
    staff_file = f'{info_folder}/{staff_files[0]}'
    staff = pd.read_excel(staff_file, skiprows=10)['Фамилия, имя, отчество']
    staff = staff.str.replace(r'\s+', ' ', regex=True)
    staff = staff.str.strip()
    staff = staff[staff.str.match(r'^\w+ \w+ \w+$', na=False)]
    return staff


# ------------------- Variables ---------------------------------------

input_folder = './data/input'
output_folder = './data/output'
info_folder = './data/info'
output_basic_file_name = 'result'
datetime_appendix = datetime.now().strftime("%d-%m-%Y_%Hh%Mm%Ss")
final_file_name = f'СКУД_АиПМБРЗ_{datetime_appendix}.xlsx'
border = '*' * 100 + '\n'  # Borders of log file parts somewhere

# ------------------- Initiate log file -------------------------------

logfile = f'{output_folder}/log_{datetime_appendix}.txt'
writelog('Создан лог-файл.')

# ------------------- Get staff names ----------------------------------

staff_files = [f for f in os.listdir(info_folder) if re.match(r'.*табель.*\.xlsx$', f, re.I)]
if not staff_files:
    writelog('!!! Нет табеля в папке data/info. Выполнение завершится !!!')
    raise Exception('No Табель file!')
staff_file = f'{info_folder}/{staff_files[0]}'
staff = get_staff_from_file(staff_file)
writelog(f'Прочитан файл табеля: {staff_files[0]}')

# ----------------- make table from pdf files --------------------------------
# There could be more than one file, need to process all of them and concatenate

pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]
if len(pdf_files) > 0:
    writelog(f'Обнаружено pdf-файлов: {len(pdf_files)}. Файлы:\n{pdf_files}')
else:
    writelog('!!! pdf-файлов не обнаружено !!!')
    raise Exception('No pdf files in input folder!')

frames = []
for file in pdf_files:
    frame = make_df_from_pdf(f'{input_folder}/{file}')
    frames.append(frame)
workframe = pd.concat(frames)
writelog('Создана единая DataFrame из pdf-файлов.')

# ---------------- check present and absent staff in the workframe ---------------

present_staff = staff[staff.isin(workframe['ФИО'])]
absent_staff = staff[~staff.isin(workframe['ФИО'])]
present_staff_string = '\n\t' + '\n\t'.join(sorted(present_staff))
absent_staff_string = '\n\t' + '\n\t'.join(sorted(absent_staff))
writelog(f'Присутствующие из табеля в списке: {present_staff_string}')
writelog(f'Отсутствующие: {absent_staff_string}')


# -------- fill dataframe with all dates (every day) from min to max date ----------------
min_date, max_date = workframe['ДАТА'].min(), workframe['ДАТА'].max()
date_range = pd.date_range(min_date, max_date)
date_range = pd.DataFrame({'ДАТА': date_range})
workframe = workframe.groupby('ФИО').apply(lambda x: x.merge(date_range, how='right', on='ДАТА'))
workframe = workframe.drop(columns=['ФИО']).reset_index(level=0).reset_index(drop=True)

# ------- Make appropriate form and data format ----------------------------------------

workframe['ЧАСЫ'] = workframe['УХОД'] - workframe['ПРИХОД'] + pd.Timestamp('1900-01-01')
workframe['ДАТА'] = workframe['ДАТА'].dt.date.apply(lambda x: x.strftime('%d.%m.%Y'))
workframe[['УХОД', 'ПРИХОД', 'ЧАСЫ']] = workframe[['УХОД', 'ПРИХОД', 'ЧАСЫ']].applymap(lambda x: x.time(), na_action='ignore')
workframe = workframe.pivot(columns='ФИО', index='ДАТА', values=['ДАТА', 'ПРИХОД', 'УХОД', 'ЧАСЫ']).reorder_levels([1, 0], axis=1)
workframe.sort_index(level=0, axis=1, inplace=True)
workframe.reset_index(drop=True, inplace=True)
workframe.columns.names = (None, None)
workframe.index = pd.RangeIndex(start=1, stop=len(workframe) + 1, step=1)
writelog('Таблица переформатирована')
workframe.to_excel(f'{output_folder}/{final_file_name}')
writelog('Таблица сохранена в .xlsx-файл')

# %%
