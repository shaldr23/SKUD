# %%
# Файл excel следует пересохранить в другой, т.к. он только для чтения
import os
import subprocess
import sys
import pandas as pd
import numpy as np
import re
from datetime import datetime


def writelog(text):
    """
    Function to securely write information into log file (logfilename)
    """
    with open(logfile, 'a') as f:
        f.write(text + '\n')


input_folder = './data/input'
output_folder = './data/output'
info_folder = './data/info'
output_basic_file_name = 'result'
datetime_appendix = datetime.now().strftime("%d-%m-%Y_%Hh%Mm%Ss")
final_file_name = f'СКУД_АиПМБРЗ_{datetime_appendix}.xlsx'
FILL_TIME = False  # fill time when either УХОД or ПРИХОД is absent (but not both)
USE_STAFF_FILE = True
std_entry_time = datetime.strptime('09:30', '%H:%M')
std_exit_time = datetime.strptime('18:00', '%H:%M')
timetable_file = 'Расписание.xlsx'
border = '*' * 100 + '\n'  # Borders of log file parts somewhere
# Log file creation
logfile = f'{output_folder}/log_{datetime_appendix}.txt'
writelog('Создан лог-файл.')

# ----------------- convert pdf files into csv --------------------------------

# There could be more than one file, need to process all of them.
pdf_files = [f for f in os.listdir(input_folder) if f.endswith('.pdf')]
if len(pdf_files) > 0:
    writelog(f'Обнаружено pdf-файлов: {len(pdf_files)}. Файлы:\n{pdf_files}')
else:
    writelog('!!! pdf-файлов не обнаружено !!!')
csv_files = []
 
for num, file in enumerate(pdf_files, start=1):
    text_file = f'{output_folder}/{output_basic_file_name}_{num}.txt'
    csv_file = f'{output_folder}/{output_basic_file_name}_{num}.csv'
    csv_files.append(csv_file)
    # !!! NEED TO OUTPUT SUBPROCESS STDOUT/STDERR
    pdftotext_process = subprocess.run(f'pdftotext.exe -table -enc UTF-8 "{input_folder}/{file}" "{text_file}"',
                                       shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                       encoding=sys.stdout.encoding)
    writelog(f'\n{border}{file}  ->  txt. Вывод pdftotext.exe:\n{pdftotext_process.stdout}\nОшибки:\n{pdftotext_process.stderr}\n{border}')
    tse_process = subprocess.run([sys.executable, 'TimeSheetExtractor2.py', text_file, csv_file],
                                 stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                 encoding=sys.stdout.encoding)
    writelog(f'\n{border}{file}  ->  ...  ->  csv. Вывод TimeSheetExtractor2.py:\n{tse_process.stdout}\nОшибки:\n{tse_process.stderr}\n{border}')

# ----- process staff (Табель) file -----------------------------

if USE_STAFF_FILE:
    staff_files = [f for f in os.listdir(info_folder) if re.match(r'.*Табель.*\.xlsx$', f, re.I)]
    if not staff_files:
        writelog('!!! Нет табеля в папке data/info. Выполнение завершится !!!')
        raise Exception('No Табель file!')
    staff_file = f'{info_folder}/{staff_files[0]}'
    staff = pd.read_excel(staff_file, skiprows=10)['Фамилия, имя, отчество']
    staff = staff.str.replace(r'\s+', ' ')
    staff = staff.str.strip()
    staff = staff[staff.str.match(r'^\w+ \w+ \w+$', na=False)]

# --------- read all csv files into one DataFrame object ----------------------

frames = []
if csv_files:
    writelog('Обработка полученных csv-файлов.')
for file in csv_files:
    try:
        frame = pd.read_csv(file, delimiter=';', encoding='utf8')
    except UnicodeDecodeError:
        frame = pd.read_csv(file, delimiter=';', encoding='cp1251')
    frames.append(frame)

if frames:
    frame = pd.concat(frames)
    frame = frame[['ФИО', 'ДАТА', 'ПРИХОД', 'УХОД']]
    frame = frame.applymap(lambda x: str(x).strip(), na_action='ignore')
    frame = frame.applymap(lambda x: np.nan if not x else x)
    frame['ДАТА'] = pd.to_datetime(frame['ДАТА'], format='%d.%m.%Y')
    frame[['ПРИХОД', 'УХОД']] = frame[['ПРИХОД', 'УХОД']].apply(lambda x: pd.to_datetime(x, format='%H:%M:%S'))
    writelog('Создана единая DataFrame из csv-файлов.')

# ---------- read xlsx files and make the same DataFrame form -----------------

xlsx_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]
if len(xlsx_files) > 0:
    writelog(f'Обнаружено xlsx-файлов: {len(xlsx_files)}. Файлы:\n{xlsx_files}')
else:
    writelog('!!! xlsx-файлов не обнаружено !!!')
frames2 = []
for file in xlsx_files:
    frame2 = pd.read_excel(f'{input_folder}/{file}', skiprows=1,
                           header=[0, 1], index_col=0)
    if frame2.empty:
        writelog(f'!!! Нечитаемый файл: {file}. Следует пересохранить его. Программа завершается !!!')
        raise Exception(f"xlsx file {file} is not readable, resave it")
    frames2.append(frame2)
    writelog(f'Файл xlsx считан: {file}')

if frames2:
    frame2 = pd.concat(frames2)
    frame2 = frame2.applymap(lambda x: str(x).strip(), na_action='ignore')
    frame2 = frame2.applymap(lambda x: np.nan if not x else x)
    cols_to_delete = [col for col in frame2.columns.levels[0] if 'unnamed' in col.lower()]
    frame2.drop(columns=cols_to_delete, level=0, inplace=True)
    frame2 = frame2.stack(0)
    frame2.columns.name = None
    frame2.reset_index(inplace=True)
    frame2.rename(columns={'level_0': 'ФИО',
                           'level_1': 'ДАТА',
                           'Entry': 'ПРИХОД',
                           'Exit': 'УХОД'},
                  inplace=True)
    frame2 = frame2[['ФИО', 'ДАТА', 'ПРИХОД', 'УХОД']]

    # filter names according to 'Табель' file
    staff_string = '\n\t' + '\n\t'.join(sorted(staff))
    writelog(f'Взяты имена из Табеля. Фильтрация таблицы из xlsx-файла проводится по {len(staff)} именам:{staff_string}')
    frame2 = frame2[frame2['ФИО'].isin(staff)]
    frame2['ДАТА'] = pd.to_datetime(frame2['ДАТА'], format='%d.%m.%y')
    frame2[['ПРИХОД', 'УХОД']] = frame2[['ПРИХОД', 'УХОД']].apply(lambda x: pd.to_datetime(x, format='%H:%M'))
    writelog('Создана единая DataFrame из xlsx-файлов.')

# ---------- union frames or treat the only one existing ------------------------------- 

# workframe will be either frame, or frame2, or union
if frames and frames2:
    merged = frame.merge(frame2, on=['ФИО', 'ДАТА'], how='outer', suffixes=('_1', '_2'))
    merged['ПРИХОД'] = merged[['ПРИХОД_1', 'ПРИХОД_2']].min(axis=1)
    merged['УХОД'] = merged[['УХОД_1', 'УХОД_2']].max(axis=1)
    to_drop = [col for col in merged.columns if re.match(r'^\w+_\d$', col)]
    merged.drop(columns=to_drop, inplace=True)
    workframe = merged
    writelog('DataFrames от данных из pdf и xlsx объединены.')
elif frames and not frames2:
    workframe = frame
elif frames2 and not frames:
    workframe = frame2

if USE_STAFF_FILE:
    # check present and absent staff in the workframe
    present_staff = staff[staff.isin(workframe['ФИО'])]
    absent_staff = staff[~staff.isin(workframe['ФИО'])]
    present_staff_string = '\n\t' + '\n\t'.join(sorted(present_staff))
    absent_staff_string = '\n\t' + '\n\t'.join(sorted(absent_staff))
    writelog(f'Присутствующие из табеля в списке: {present_staff_string}')
    writelog(f'Отсутствующие: {absent_staff_string}')

# fill dataframe with all dates (every day) from min to max date 
min_date, max_date = workframe['ДАТА'].min(), workframe['ДАТА'].max()
date_range = pd.date_range(min_date, max_date)
date_range = pd.DataFrame({'ДАТА': date_range})
workframe = workframe.groupby('ФИО').apply(lambda x: x.merge(date_range, how='right', on='ДАТА'))
workframe = workframe.drop(columns=['ФИО']).reset_index(level=0).reset_index(drop=True)
# fill default time according to personal timetable when either УХОД or ПРИХОД is absent (but not both)
if FILL_TIME:
    to_fill_bool = (workframe['ПРИХОД'].isna() & ~workframe['УХОД'].isna()) | (~workframe['ПРИХОД'].isna() & workframe['УХОД'].isna())
    to_fill = workframe[to_fill_bool]
    if not to_fill.empty:
        timetable = pd.read_excel(f'{info_folder}/{timetable_file}')
        timetable.dropna(inplace=True, subset=['ФИО'])
        timetable[['ПРИХОД', 'УХОД']] = timetable[['ПРИХОД', 'УХОД']].applymap(lambda x: datetime.strptime(x.isoformat(), '%H:%M:%S'))
        writelog('Обнаружены недостающие данные ПРИХОД/УХОД (есть один показатель, но нет другого):\n' + to_fill.to_csv(sep="\t", na_rep='NA'))
        fillmerged = to_fill.merge(timetable, how='left', on='ФИО')
        fillmerged.index = to_fill.index
        to_fill['ПРИХОД'] = to_fill['ПРИХОД'].fillna(fillmerged['ПРИХОД_y'])
        to_fill['ПРИХОД'] = to_fill['ПРИХОД'].fillna(std_entry_time)
        to_fill['УХОД'] = to_fill['УХОД'].fillna(fillmerged['УХОД_y'])
        to_fill['УХОД'] = to_fill['УХОД'].fillna(std_exit_time)
        workframe[to_fill_bool] = to_fill
        writelog('Данные заполнены следующим образом:\n' + to_fill.to_csv(sep="\t", na_rep='NA'))

# ------- Make appropriate form and data format --------------------------------------------------------------------

workframe['ЧАСЫ'] = workframe['УХОД'] - workframe['ПРИХОД'] + pd.Timestamp('1900-01-01')
workframe['ДАТА'] = workframe['ДАТА'].dt.date.apply(lambda x: x.strftime('%d.%m.%Y'))
workframe[['УХОД', 'ПРИХОД', 'ЧАСЫ']] = workframe[['УХОД', 'ПРИХОД', 'ЧАСЫ']].applymap(lambda x: x.time(), na_action='ignore')
workframe = workframe.pivot(columns='ФИО', index='ДАТА', values=['ДАТА', 'ПРИХОД', 'УХОД', 'ЧАСЫ']).reorder_levels([1, 0], axis=1)
workframe.sort_index(level=0, axis=1, inplace=True)
workframe.reset_index(drop=True, inplace=True)
workframe.columns.names = (None, None)
writelog('Таблица переформатирована')
workframe.to_excel(f'{output_folder}/{final_file_name}')
writelog('Таблица сохранена в .xlsx-файл')

# %%
