import pandas as pd
from openpyxl import Workbook
from datetime import datetime, date, time
from my_functions import diagram


#Переменные для управления пользователем
excel_file_path = 'Performance Management-History Query-LTE_Main_KPIs_ITBBU-bkhayr-20230613094954.xlsx'
date_time_str = '2023-06-12 00:00:00'
show_compare_date=True
report_path_dir = 'C:\\Users\\User\\Desktop\\зн\\Report\\'
show_peaks=True
show_valleys=True
peaks_color='green'
valleys_color='blue'

# Загрузите данные из файла Excel
df = pd.read_excel(excel_file_path)
# Get the index of the 'Subnetwork' column, нижний цикл будет срабатывать корректно в этом случае, если не добавлять пустой столбец тогда не будет отабражаться плот первого столбца после Subnetork
index = df.columns.get_loc('Subnetwork')
# Insert an empty column after the 'Subnetwork' column
df.insert(index + 1, 'New Column', '')



# Преобразуйте названия столбцов в нижний регистр
columns = [column.lower() for column in df.columns]
print(columns)
# Проверьте наличие столбца 'subnetwork'
if 'Subnetwork' not in df.columns:
    print('Ошибка: Столбец с названием Subnetwork не найден')
elif 'Start Time' not in df.columns:
	print('Ошибка: Столбец с названием Start Time не найден')
else:
    # Получите индекс столбца 'subnetwork'
    index = df.columns.get_loc('Subnetwork')
    # Выведите общее количество столбцов
    print('Общее количество столбцов:', len(df.columns))
    # Выведите общее количество столбцов после столбца 'Subnetwork'
    column_count = len(df.columns[index+1:])
    print('Общее количество столбцов после столбца Subnetwork: {}'.format(column_count))
    # Создайте список ячеек для размещения графиков
    cell_names = ['A1', 'N1']
    # Создайте рабочую книгу и лист
    workbook = Workbook()
    sheet = workbook.active
    for i in range(2, column_count+1):
        if i % 2 == 0:
            cell_names.append('A'+str(20*(i//2)))
        else:
            cell_names.append('N'+str(20*(i//2)))
        # Выберите нужные столбцы для отображения на графике
        kpi1 = df[df.columns[index+i]]
        time = df['Start Time']
        # Выведите график в соответствующую ячейку, вызываем функцию диаграм
        diagram(time, kpi1, cell_names[-1], date_time_str, sheet, report_path_dir, show_peaks, show_valleys, peaks_color, valleys_color, show_compare_date)
    #время для создания уникального файла
    now = datetime.now()
    report_name = 'Report_time_' + now.strftime('%Y-%m-%d_%H-%M-%S') + '.xlsx'
    report_path = report_path_dir + report_name
    workbook.save(report_path)
    print('Finished: '+ str(now))


