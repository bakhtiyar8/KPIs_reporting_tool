import pandas as pd
import matplotlib.pyplot as plt
from scipy.signal import find_peaks
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from datetime import datetime, date, time

def diagram(time, kpi1, subnetwork_values, cell_name, date_time_str, sheet, report_path_dir, show_peaks, show_valleys, peaks_color='red', valleys_color='blue', show_compare_date=False):
	# Установите минимальное расстояние между пиками
	min_distance = 10
	# Найдите индексы пиковых значений в данных
	peaks, _ = find_peaks(kpi1, distance=min_distance)
	# Инвертируйте данные
	inverted_kpi1 = 1/kpi1
	# Найдите индексы низких значений в данных
	valleys, _ = find_peaks(inverted_kpi1, distance=min_distance)
	# Создайте фигуру с одним графиком
	fig, ax = plt.subplots(figsize=(10.3, 3))
	ax.plot(time, kpi1)
	ax.set_title(str(kpi1.name))
	# Добавьте вертикальную линию на график
	if show_compare_date:
		ax.axvline(x=date_time_str, color='red', linestyle='--')
	# Измените интервал отображения времени на оси X
	ax.set_xticks(time[::3])
	# Измените размер шрифта меток времени на оси X
	ax.tick_params(axis='x', labelsize=6.5)
	ax.tick_params(axis='y', labelsize=8)
	# Поверните метки времени на оси X
	plt.setp(ax.get_xticklabels(), rotation=90)
	# Отобразите пиковые значения на графике
	if show_peaks:
		ax.plot(time[peaks], kpi1[peaks], 'o', color=peaks_color)
	# Отобразите низкие значения на графике
	if show_valleys:
		ax.plot(time[valleys], kpi1[valleys], 'o', color=valleys_color)
	# Set the desired font size
	# Set the y-axis limits
	fontsize = 8
	if show_peaks:
		# Add KPI values next to peak values
		for x, y in zip(time[peaks], kpi1[peaks]):
			ax.annotate(f'{y:.2f}', xy=(x, y), xytext=(0, -15), textcoords='offset points', ha='center', va='bottom', fontsize=fontsize)
	if show_valleys:
		# Add KPI values next to low values
		for x, y in zip(time[valleys], kpi1[valleys]):
			ax.annotate(f'{y:.2f}', xy=(x, y), xytext=(0, 15), textcoords='offset points', ha='center', va='top', fontsize=fontsize)
	# Создайте список меток для легенды
	labels = ['Legend:']
	if show_compare_date:
		labels.append('Operation Time')
	if show_peaks:
		labels.append('Peaks')
	if show_valleys:
		labels.append('Valleys')
	# Добавьте легенду в нижней части графика
	#ax.legend(labels=labels, loc='lower center', bbox_to_anchor=(0.5, -0.6), ncol=len(labels), fontsize=7)
	# Добавьте линии в плоты в зависимости от количества названий в колонке "subnetwork_values"
	subnetworks = subnetwork_values.unique()
	for i, sub in enumerate(subnetworks):
		sub_time = time[subnetwork_values == sub]
		sub_kpi1 = kpi1[subnetwork_values == sub]
		ax.plot(sub_time, sub_kpi1, label=sub)
	# Добавьте названия подсетей в список меток легенды
	labels += list(subnetworks)
	# Обновите легенду с названиями из "subnetwork_values"
	ax.legend(labels=labels, loc='lower center', bbox_to_anchor=(0.5, -0.6), ncol=len(labels), fontsize=7)
	#строка названия изображения
	pic_name = 'KPI_{}.png'.format(kpi1.name.replace('/', '_'))
	report_path = report_path_dir+ "/" + pic_name
	plt.savefig(report_path, bbox_inches='tight')
	plt.close(fig)
	# Сохраните фигуру в виде изображения
	img = Image(report_path)
	#cell_name пример A1
	sheet.add_image(img, cell_name)