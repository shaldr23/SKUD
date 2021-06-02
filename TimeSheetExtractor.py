# TimeSheetExtractor, v2020.1.2.5, Турубар Д.С.
# Извлечение графика прихода-ухода из текстового представления PDF-файла
#
# Для работы необходим pdftotext.exe из пакета XpdfReader
# Схема работы:
# - извлечение текста из PDF-файла при помощи pdftotext.exe
# - обработка полученного текста
# - запись результата в формате CSV-таблицы (можно открыть табличным процессором, типа Excel)
#
# Для запуска:
#_convert.bat путь_к_pdf_файлу
# В этой же папке должны находиться файлы: pdftotext.exe, TimeSheetExtractor.py, установлен Python
#--
import sys
import os.path
import re

# проверки параметров командной строки
# кол-во  аргументов
if len(sys.argv) != 3:
    print(f"""Ошибка: количество аргументов командной строки = {len(sys.argv)-1}
Ожидается: 2
    1: Путь к входному файлу. Например: time_sheet__pdf-txt.txt
    2: Путь к выходному файлу. Например: time_sheet__txt-csv.csv
""")
    exit()

# входной файл
if not os.path.exists( sys.argv[1] ):
    print(f"Ошибка: входной файл не найден: { sys.argv[1] }")
    exit()


# основные переменные
inp_fname = sys.argv[1]
out_fname = sys.argv[2]
lines = []

# 1 - загрузка данных
print(f'Загрузка файла: {inp_fname}')
with open(inp_fname, 'r', encoding='utf8') as f:
    lines = f.readlines()
print(f'  строк: {len(lines)}')

# 2 - нормализация загруженных данных
print('\nУдаление пустых строк')
lines = [l for l in lines if l.strip()]
print(f'  строк: {len(lines)}')

# 3 - извлечение данных
print(f'\nИзвлечение данных:')


# открытие файла на запись
csv = open(out_fname, 'w')
csv.write('ФИО;ДАТА;ПРИХОД;УХОД;\n')

# инициализация
fio    = []
date_action = { }

# обход с 1 записи (работа с 2 строками, текущая и предыдущая)
count = 0
for i in range(1, len(lines) ): 
    # строки
    l_prev = lines[i-1].lstrip() # предыдущая
    l_cur = lines[i].lstrip()	 # текущая

    if re.match(r'^\s*фамилия', l_prev):

        if len(fio) == 3:
            # запись в файл
            for date, v in date_action.items():
                csv.write(f'{fio[0]} {fio[1]} {fio[2]}; { date }; { v["login"][0] if len(v["login"]) > 0 else "" }; { v["logout"][-1] if len(v["logout"]) > 0 else "" } \n');
        
        #очистка текущей записи
        fio.clear()
        date_action.clear()

        # добавление фамилии
        fio.append( re.search(r'^\s*(\w+)', l_cur)[1] )

    if re.match(r'^\s*имя', l_prev):
        # добавление имени
        fio.append( re.search(r'^\s*(\w+)', l_cur)[1] )

    if re.match(r'^\s*отчество', l_prev):
        # добавление отчества
        fio.append( re.search(r'^\s*(\w+)', l_cur)[1] )
        count += 1
        print(f"{count:03}: {fio[0]} {fio[1]} {fio[2]} ")

    if len(fio) == 3:
        
        l_cur_action = re.match(r'^(\d{2}\.\d{2}\.\d{4})\s*(\d{2}\:\d{2}\:\d{2})\s*(Вы?ход)', l_cur)

        if l_cur_action:
            l_cur_date = l_cur_action[1]
            l_cur_time = l_cur_action[2]        
            l_cur_type = l_cur_action[3]

            if 'Вход' in l_cur_type:
                
                if l_cur_date in date_action:
                    date_action[ l_cur_date ]['login'].append( l_cur_time )
                else:
                    date_action[ l_cur_date ] = {}
                    date_action[ l_cur_date ]['login']  = []
                    date_action[ l_cur_date ]['logout'] = []                    
                    
                    date_action[ l_cur_date ]['login'].append( l_cur_time )
                    

            if 'Выход' in l_cur_type:

                if l_cur_date in date_action:
                    date_action[ l_cur_date ]['logout'].append( l_cur_time )
                else:
                    date_action[ l_cur_date ] = {}
                    date_action[ l_cur_date ]['login']  = []
                    date_action[ l_cur_date ]['logout'] = []                    
                    
                    date_action[ l_cur_date ]['logout'].append( l_cur_time )
                
        
# запись файла
csv.close()
