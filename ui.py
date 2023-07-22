import tkinter as tk
from tkinter import filedialog
import pandas as pd
from openpyxl import Workbook
from datetime import datetime, date, time
from my_functions import diagram
# Import the ttk module from tkinter
from tkinter import ttk
from tkinter import Label

def select_excel_file():
    # Open a file dialog to select the excel file
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    # Update the input field with the selected file path
    excel_file_path_entry.delete(0, tk.END)
    excel_file_path_entry.insert(0, excel_file_path)

    # Load the data from the selected file
    df = pd.read_excel(excel_file_path)
    # Convert the 'Start Time' column to a string format
    df['Start Time'] = df['Start Time'].astype(str)
    # Get the unique values from the 'Start Time' column
    start_times = list(df['Start Time'].unique())
    # Clear the date_time_str_entry field
    date_time_str_entry.delete(0, tk.END)
    # Create a combobox widget to allow the user to select a date date_time_str_var date_time_str_entry
    date_time_str_combobox = ttk.Combobox(root, textvariable=date_time_str_var, values=start_times)
    date_time_str_combobox.grid(row=1, column=2)
    def on_select(event):
        selected_value = date_time_str_combobox.get()
        date_time_str_entry.delete(0, tk.END)
        date_time_str_entry.insert(0, selected_value)
    date_time_str_combobox.bind("<<ComboboxSelected>>", on_select)

def select_report_dir():
    # Open a directory dialog to select the report directory
    report_path_dir = filedialog.askdirectory()
    report_path_dir = report_path_dir + '/'
    # Update the input field with the selected directory path
    report_path_dir_entry.delete(0, tk.END)
    report_path_dir_entry.insert(0, report_path_dir)

def run_program():
    # Get the values from the input fields
    excel_file_path = excel_file_path_entry.get()
    date_time_str = date_time_str_entry.get()
    print('Run '+date_time_str)
    show_compare_date = show_compare_date_var.get()
    report_path_dir = report_path_dir_entry.get()
    show_peaks = show_peaks_var.get()
    show_valleys = show_valleys_var.get()
    peaks_color = peaks_color_entry.get()
    valleys_color = valleys_color_entry.get()

    # Run your program here using the values from the input fields
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
            subnetwork_values = df['Subnetwork']
            # Выведите график в соответствующую ячейку, вызываем функцию диаграм
            diagram(time, kpi1, subnetwork_values, cell_names[-1], date_time_str, sheet, report_path_dir, show_peaks, show_valleys, peaks_color, valleys_color, show_compare_date)
        #время для создания уникального файла
        now = datetime.now()
        report_name = 'Report_time_' + now.strftime('%Y-%m-%d_%H-%M-%S') + '.xlsx'
        report_path = report_path_dir + report_name
        workbook.save(report_path)
        print('Finished: '+ str(now))



# Create the main window
root = tk.Tk()
# Установите размер окна
root.geometry("1000x300")
# Создание главного фрейма
frame = tk.Frame(root)
frame.grid()
# Добавление своих инициалов
initials_label = tk.Label(root, text="Software made by Khayrullaev Bakhtiyar, code: https://github.com/bakhtiyar8/KPIs_reporting_tool")
initials_label.grid(row=25, column=2, sticky="se")  # Размещение в последней строке и последнем столбце
# Создайте виджет Label с комментарием или сообщением
comment_label = tk.Label(root, text="In the excel file(*.xlsx), columns with the names Start Time and Subnetwork are required!")
# Разместите виджет Label в окне
comment_label.grid(row=0, column=4)
# Create a string variable to store the selected date
date_time_str_var = tk.StringVar()

# Create the input fields and labels
excel_file_path_label = tk.Label(root, text="Excel File Path")
excel_file_path_label.grid(row=1, column=0)
excel_file_path_entry = tk.Entry(root)
excel_file_path_button = tk.Button(root, text="Browse", command=select_excel_file)
date_time_str_label = tk.Label(root, text="Operation Date Time")
date_time_str_label.grid(row=2, column=0)
date_time_str_entry = tk.Entry(root)
show_compare_date_label = tk.Label(root, text="Show Compare Date")
show_compare_date_label.grid(row=3, column=0)
show_compare_date_var = tk.BooleanVar()
show_compare_date_checkbutton = tk.Checkbutton(root, variable=show_compare_date_var)
report_path_dir_label = tk.Label(root, text="Report Path Dir")
report_path_dir_label.grid(row=4, column=0)
report_path_dir_entry = tk.Entry(root)
report_path_dir_button = tk.Button(root, text="Browse", command=select_report_dir)
show_peaks_label = tk.Label(root, text="Show Peaks")
show_peaks_label.grid(row=5, column=0)
show_peaks_var = tk.BooleanVar()
show_peaks_checkbutton = tk.Checkbutton(root, variable=show_peaks_var)
show_valleys_label = tk.Label(root, text="Show Valleys")
show_valleys_label.grid(row=6, column=0)
show_valleys_var = tk.BooleanVar()
show_valleys_checkbutton = tk.Checkbutton(root, variable=show_valleys_var)
peaks_color_label = tk.Label(root, text="Peaks Color")
peaks_color_label.grid(row=7, column=0)
peaks_color_entry = tk.Entry(root)
valleys_color_label = tk.Label(root, text="Valleys Color")
valleys_color_label.grid(row=8, column=0)
valleys_color_entry = tk.Entry(root)

# Create the run button
run_button = tk.Button(root, text="Run", command=run_program)



# Layout the widgets using the grid layout manager
excel_file_path_label.grid(row=0, column=0)
excel_file_path_entry.grid(row=0, column=1)
excel_file_path_button.grid(row=0, column=2)
date_time_str_label.grid(row=1, column=0)
date_time_str_entry.grid(row=1, column=1)
show_compare_date_label.grid(row=2, column=0)
show_compare_date_checkbutton.grid(row=2, column=1)
report_path_dir_label.grid(row=3, column=0)
report_path_dir_entry.grid(row=3, column=1)
report_path_dir_button.grid(row=3, column=2)
show_peaks_label.grid(row=4, column=0)
show_peaks_checkbutton.grid(row=4, column=1)
show_valleys_label.grid(row=5, column=0)
show_valleys_checkbutton.grid(row=5, column=1)
peaks_color_label.grid(row=6, column=0)
peaks_color_entry.grid(row=6, column=1)
valleys_color_label.grid(row=7, column=0)
valleys_color_entry.grid(row=7, column=1)

run_button.grid(row=8,columnspan=2)

# Run the main loop
root.mainloop()
