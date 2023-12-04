from functions import gpx_to_excel, proj4_to_dict, gpx_to_msk_txt, gpx_to_kml, kml_to_gpx, check_extension, total_table, txt_msk_to_exl, autocad_msk_to_exl, xls_to_kml, xls_to_gpx
import sys, os
import pickle
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QApplication, QDesktopWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView, QFileDialog, QTableWidgetSelectionRange, 
	QMessageBox, QCheckBox, QComboBox, QDialog, QRadioButton, QSpinBox, QLabel, QSizePolicy)
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon


list_of_param_alt = ['ord', 'name', 'lat', 'lon', 'time', 'date', 'cmt'] 
headers_alt = ['Номер', 'Имя', 'Северная широта', 'Восточная долгота', 'Время', 'Дата', 'Комментарий'] 

dict_proj4_data = proj4_to_dict('proj4.txt')

try:
	with open('data_atr.pickle', 'rb') as file:
		atr_dict = pickle.load(file)
except:
	atr_dict = {}

	atr_dict['exl'] = False
	atr_dict['txt'] = False
	atr_dict['srt'] = False
	atr_dict['sp'] = 10
	atr_dict['err'] = 20
	atr_dict['xy'] = 'XY'
	atr_dict['path'] = os.getcwd()

	with open('data_atr.pickle', 'wb') as file:
		pickle.dump(atr_dict, file)


class AdvTable(QTableWidget):
	def __init__(self, parent):
		super().__init__(parent)

		self.main = parent

		self.setAcceptDrops(True)

	def dragEnterEvent(self, event):
		data = event.mimeData()
		urls = data.urls()
		if urls and urls[0].scheme() == 'file':
			event.acceptProposedAction()

	def dragMoveEvent(self, event):
		data = event.mimeData()
		urls = data.urls()
		if urls and urls[0].scheme() == 'file':
			event.acceptProposedAction()

	def dropEvent(self, event):
		is_true = True
		data = event.mimeData()
		urls = data.urls()
		address_list = []
		if urls and urls[0].scheme() == 'file':

			for filepath in urls:
				abs_path = str(filepath.path())[1:]
				extension = os.path.splitext(abs_path)[-1]

				if extension not in ['.gpx', '.kml', '.xls', '.txt']:
					is_true = False
					break

			if is_true:
				for filepath in urls:
					address_list.append(str(filepath.path())[1:])

				self.main.total.extend(address_list)
				self.main.from_list_to_table(self.main.total)
			else:
				QMessageBox.information(self, 'Сообщение', 'Неверный формат', QMessageBox.Ok)


class Dialog(QDialog):
	def __init__(self, root):
		super().__init__(root)
		self.setWindowTitle('Параметры')

		self.initExl = atr_dict['exl']
		self.initTxt = atr_dict['txt']
		self.initSp = atr_dict['sp']
		self.initErr = atr_dict['err']
		self.initXY = atr_dict['xy']
		self.initSrt = atr_dict['srt']

		self.main = root

		self.setMaximumSize(220, 180) 

		self.label = QLabel('Дробная часть')
		self.label2 = QLabel('Порядок вывода координат в МСК:')
		self.label3 = QLabel('Относительная ошибка (%)')

		self.btnOk = QPushButton('Ok')
		self.btnCnsl = QPushButton('Отмена')
		
		self.checkExl = QCheckBox('Добавить комментарии в вывод Excel')
		self.checkMsk = QCheckBox('Добавить имя в вывод МСК')
		self.checkSrt = QCheckBox('Сортировка данных в выводе Excel')

		# spinboxes
		self.sp = QSpinBox()

		self.sp.setRange(1, 10)
		self.sp.setValue(10)

		self.sp.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Expanding)
		self.sp.setMaximumWidth(50)
		self.sp.setMaximumHeight(17)

		self.err = QSpinBox()
		
		self.err.setRange(1, 100)
		self.err.setValue(20)

		self.err.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Expanding)
		self.err.setMaximumWidth(50)
		self.err.setMaximumHeight(17)

		# end spinboxes

		self.rad1 = QRadioButton('XY (Широта, Долгота)')
		self.rad2 = QRadioButton('YX (Долгота, Широта)')

		self.rad1.setChecked(True)

		self.dhbox1 = QHBoxLayout()
		self.dhbox2 = QHBoxLayout()
		self.dhbox3 = QHBoxLayout()
		self.dhbox4 = QHBoxLayout()
		self.dvbox = QVBoxLayout()

		self.dhbox1.addWidget(self.rad1)
		self.dhbox1.addWidget(self.rad2)
																
		self.dhbox2.addStretch(1)
		self.dhbox2.addWidget(self.btnOk)
		self.dhbox2.addWidget(self.btnCnsl)

		self.dhbox3.addWidget(self.sp)
		self.dhbox3.addWidget(self.label)

		self.dhbox4.addWidget(self.err)
		self.dhbox4.addWidget(self.label3)
		
		self.btnCnsl.clicked.connect(self.on_btnCnsl)
		self.btnOk.clicked.connect(self.on_btnOk)
		self.rad1.toggled.connect(lambda:self.on_rad(self.rad1))
		self.rad2.toggled.connect(lambda:self.on_rad(self.rad2))

		self.dvbox.addWidget(self.checkSrt)
		self.dvbox.addWidget(self.checkExl)
		self.dvbox.addWidget(self.checkMsk)
		self.dvbox.addLayout(self.dhbox3)
		self.dvbox.addLayout(self.dhbox4)
		self.dvbox.addWidget(self.label2)
		self.dvbox.addLayout(self.dhbox1)
		self.dvbox.addLayout(self.dhbox2)

		self.setLayout(self.dvbox)

		self.assignWidgets()

	def on_btnCnsl(self):
		self.checkExl.setChecked(self.initExl)
		self.checkMsk.setChecked(self.initTxt)
		self.checkSrt.setChecked(self.initSrt)
		self.sp.setValue(self.initSp)
		self.err.setValue(self.initErr)
		
		if self.initXY == 'XY':
			self.rad1.setChecked(True)
		elif self.initXY == 'YX':
			self.rad2.setChecked(True)
		
		self.close()

	def on_btnOk(self):
		self.main.exlCheck = self.checkExl.isChecked()
		self.main.txtCheck = self.checkMsk.isChecked()
		self.main.srtCheck = self.checkSrt.isChecked()
		self.main.txtParam = self.sp.value()
		self.main.errParam = self.err.value()
		self.main.xy = 'XY' if self.rad1.isChecked() else 'YX'

		self.initExl = self.main.exlCheck
		self.initTxt = self.main.txtCheck
		self.initSrt = self.main.srtCheck
		self.initSp = self.main.txtParam
		self.initErr = self.main.errParam
		self.initXY = self.main.xy

		atr_dict['exl'] = self.checkExl.isChecked()
		atr_dict['txt'] = self.checkMsk.isChecked()
		atr_dict['srt'] = self.checkSrt.isChecked()
		atr_dict['sp'] = self.sp.value()
		atr_dict['err'] = self.err.value()
		atr_dict['xy'] = 'XY' if self.rad1.isChecked() else 'YX'

		with open('data_atr.pickle', 'wb') as file:
			pickle.dump(atr_dict, file)

		self.close()

	def assignWidgets(self):
		self.checkExl.setChecked(atr_dict['exl'])
		self.checkMsk.setChecked(atr_dict['txt'])
		self.checkSrt.setChecked(atr_dict['srt'])
		self.sp.setValue(atr_dict['sp'])
		self.err.setValue(atr_dict['err'])
		if atr_dict['xy'] == "XY":
			self.rad1.setChecked(True)
		elif atr_dict['xy'] == "YX":
			self.rad2.setChecked(True)

class Wait(QDialog):
	def __init__(self, root):
		super().__init__(root)
		self.setMaximumSize(220, 180)
		self.label = QLabel("Идет конвертация данных. Пожалуйста подождите...")
		self.dvbox = QVBoxLayout()
		self.dvbox.addWidget(self.label)
		self.setLayout(self.dvbox)



class Window(QWidget):
	def __init__(self):
		super().__init__()
		self.resize(500, 400)
		self.setWindowTitle('Get Points 1.3')
		self.setWindowIcon(QIcon('icons/compass.png'))
		qr = self.frameGeometry()
		cp = QDesktopWidget().availableGeometry().center()
		qr.moveCenter(cp)
		self.move(qr.topLeft())

		self.setAcceptDrops(True)

		# Attributes
		self.current_selected_row = None
		self.total = [] # All paths to files gpx for move to exl is here

		self.exlCheck = atr_dict['exl']
		self.txtCheck = atr_dict['txt']
		self.srtCheck = atr_dict['srt']
		self.txtParam = atr_dict['sp']
		self.errParam = atr_dict['err']
		self.xy = atr_dict['xy']
		# End

		# Dialog window
		self.dialog = Dialog(self)
		self.wait = Wait(self)
		# End

		self.vbox = QVBoxLayout()
		self.hbox_1 = QHBoxLayout()
		self.hbox_2 = QHBoxLayout()
		self.hbox_3 = QHBoxLayout()

		self.table = AdvTable(self)
		
		# --buttons--
		self.btnAdd = QPushButton('Добавить')
		self.btnDel = QPushButton('Удалить')
		self.btnExl = QPushButton('В Excel')
		self.btnMsk = QPushButton('МСК')
		self.btnParam = QPushButton('Параметры')
		self.btnGpxToKml = QPushButton('GPX в KML')
		self.btnKmlToGpx = QPushButton('KML в GPX')
		self.btnMerge = QPushButton('Компоновка')
		# --end--

		# toolTips for buttons
		self.btnAdd.setToolTip("Довавить файлы. Возможны расширения: <b>'xls'</b>, <b>'gpx'</b>, <b>'kml'</b> ")
		self.btnDel.setToolTip("Удалить из списка. При нажатии клавиши <b>Delete</b> очищается весь список ")
		self.btnExl.setToolTip("Формирует таблицу <b>Excel</b> с наименованием точек и соответствующими координатами")
		self.btnMsk.setToolTip("Создает файл формата <b>txt</b> с координатами <b>МСК</b> или <b>WGS-84</b>")
		self.btnParam.setToolTip("Общие параметры программы")
		self.btnGpxToKml.setToolTip("Конвертирует <b>gpx</b> файл в <b>kml</b> файл")
		self.btnKmlToGpx.setToolTip("Конвертирует <b>kml</b> файл в <b>gpx</b> файл")
		self.btnMerge.setToolTip("Объединаяет файл <b>xls</b> с координатами точек и файл <b>xls</b> со значениями мощности дозы в единый файл")
		# End

		self.combobox = QComboBox()

		for key in dict_proj4_data.keys():
			self.combobox.addItem(key)

		self.table.setColumnCount(2)
		self.table.setRowCount(10)
		self.table.setHorizontalHeaderLabels(['Name', 'Date'])
		self.table.horizontalHeaderItem(0).setTextAlignment(Qt.AlignHCenter)
		self.table.horizontalHeaderItem(1).setTextAlignment(Qt.AlignHCenter)

		self.setStyleSheet("QTableView{ selection-background-color: rgba(255, 0, 0, 30); selection-color: black;  }")
		self.btnParam.setStyleSheet("QPushButton{ background-color: rgba(200, 0, 0, 30);  }")
	
		header = self.table.horizontalHeader()
		header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
		header.setStretchLastSection(True)

		self.hbox_1.addWidget(self.btnMerge)
		self.hbox_1.addStretch(1)
		self.hbox_1.addWidget(self.btnAdd)
		self.hbox_1.addWidget(self.btnDel)
		self.hbox_1.addWidget(self.btnGpxToKml)

		self.hbox_2.addWidget(self.btnParam)
		self.hbox_2.addStretch(1)
		self.hbox_2.addWidget(self.btnExl)
		self.hbox_2.addWidget(self.btnMsk)
		self.hbox_2.addWidget(self.btnKmlToGpx)
		
		self.hbox_3.addWidget(self.combobox)

		self.vbox.addLayout(self.hbox_3)
		self.vbox.addWidget(self.table)
		self.vbox.addLayout(self.hbox_1)
		self.vbox.addLayout(self.hbox_2)

		self.setLayout(self.vbox)

		self.btnAdd.clicked.connect(self.on_btnAdd)
		self.btnDel.clicked.connect(self.on_btnDel)
		self.btnExl.clicked.connect(self.on_btnExl)
		self.btnMsk.clicked.connect(self.on_btnMsk)
		self.btnGpxToKml.clicked.connect(self.on_btnGpxToKml)
		self.btnKmlToGpx.clicked.connect(self.on_btnKmlToGpx)
		self.btnMerge.clicked.connect(self.on_btnMerge)
		
		self.btnParam.clicked.connect(self.dialog.exec)

		self.table.cellClicked.connect(self.cell_was_clicked)
	
	# Button clicked events
	def on_btnAdd(self):
		try:
			path_f = atr_dict['path']
		except:
			path_f = os.getcwd()
		fnames = QFileDialog.getOpenFileNames(self, 'Open file', path_f, "Files (*.gpx *.kml *.xls)")[0]
		
		if fnames:
			atr_dict['path'] = os.path.dirname(fnames[0])
			with open('data_atr.pickle', 'wb') as file:
				pickle.dump(atr_dict, file)


		for fname in fnames:
			if fname not in self.total:
				self.total.append(fname)
		
		self.from_list_to_table(self.total)

	def on_btnDel(self):
		selected = set(index.row() for index in self.table.selectedIndexes())
		
		if selected:
			items = []
			
			for ind in selected:
				items.append(self.table.item(ind, 0))
			
			filenames = []

			for item in items:
				if item: filenames.append(item.text())

			temp_list = []

			for ind, file in enumerate(self.total):
				rel = file.split('/')[-1]
				if rel in filenames: temp_list.append(file)

			self.total = list(set(self.total) - set(temp_list))
			self.table.clearContents()
			self.from_list_to_table(self.total)

		
	def on_btnExl(self):
		try:
			with open('data_atr.pickle', 'rb') as file:
				atr_dict = pickle.load(file)
			
			path_f = atr_dict['path']
		except:
			path_f = os.getcwd()
		
		if self.total:
			if check_extension(".gpx", self.total) or check_extension(".kml", self.total):
				path = QFileDialog.getSaveFileName(self, 'Сохранить файл', path_f, 'XLS(*.xls)')
				if path[0] != '':
					try:
						gpx_to_excel(self.total, path[0], headers_alt, list_of_param_alt, comment=self.exlCheck, srt=self.srtCheck)
					except:
						QMessageBox.information(self, "Сообщение", "Что-то пошло не так. Проверьте формат правильность заполнения.")
					else:
						QMessageBox.information(self, 'Сообщение', 'Информация сконвертирована в файл excel', QMessageBox.Ok)
					
					atr_dict["path"] = os.path.dirname(path[0])

					with open('data_atr.pickle', 'wb') as file:
						pickle.dump(atr_dict, file)
			else:
				QMessageBox.information(self, "Сообщение", "Неверный формат файлов")


		else:
			QMessageBox.information(self, 'Сообщение', 'Пустая таблица', QMessageBox.Ok)


	def on_btnMsk(self):
		try:
			with open('data_atr.pickle', 'rb') as file:
				atr_dict = pickle.load(file)

			path_f = atr_dict['path']
		except:
			path_f = os.getcwd()
		
		zone_key = self.combobox.currentText()
		
		if self.total:
			
			if check_extension('.gpx', self.total) or check_extension('.kml', self.total):
				path = QFileDialog.getSaveFileName(self, 'Сохранить файл', path_f, 'TXT(*.txt)')
				if path[0] != '':
					if self.txtCheck:
						gpx_to_msk_txt(self.total, path[0], zone_key, self.txtParam, self.xy, add_name=True)
					else:
						gpx_to_msk_txt(self.total, path[0], zone_key, self.txtParam, self.xy, add_name=False)
					
					text_for_msg = 'Информация конвертирована в ' + zone_key
					QMessageBox.information(self, 'Сообщение', text_for_msg, QMessageBox.Ok)

					atr_dict["path"] = os.path.dirname(path[0])

					with open('data_atr.pickle', 'wb') as file:
						pickle.dump(atr_dict, file)
			
			elif check_extension('.txt', self.total):
				path = QFileDialog.getSaveFileName(self, "Сохранить файл", path_f, 'XLS(*.xls)')
				
				if path[0] != '':
					response = txt_msk_to_exl(self.total, path[0], zone_key)
					
					if response == -1:
						QMessageBox.information(self, 'Сообщение', "Ошибка конвертации. Проверьте правильность заполнения файла", QMessageBox.Ok)
					elif response == 1:
						text_for_msg = 'Информация конвертирована в ' + zone_key
						QMessageBox.information(self, 'Сообщение', text_for_msg, QMessageBox.Ok)

						atr_dict["path"] = os.path.dirname(path[0])

						with open('data_atr.pickle', 'wb') as file:
							pickle.dump(atr_dict, file)

			elif check_extension('.xls', self.total):
				path = QFileDialog.getSaveFileName(self, "Сохранить файл", path_f, 'XLS(*.xls)')
				
				if path[0] != '':
					response = autocad_msk_to_exl(self.total, path[0], zone_key, self.txtParam)
					if response == -1:
						QMessageBox.information(self, 'Сообщение', "Ошибка конвертации. Проверьте правильность заполнения файла", QMessageBox.Ok)
					elif response == 1:
						text_for_msg = 'Информация конвертирована в ' + zone_key
						QMessageBox.information(self, 'Сообщение', text_for_msg, QMessageBox.Ok)

						atr_dict["path"] = os.path.dirname(path[0])

						with open('data_atr.pickle', 'wb') as file:
							pickle.dump(atr_dict, file)
			
			else:
				QMessageBox.information(self, 'Сообщение', 'Неверный формат файлов или файлы разного формата', QMessageBox.Ok)
		else:
			QMessageBox.information(self, 'Сообщение', 'Пустая таблица', QMessageBox.Ok)

	def on_btnGpxToKml(self):
		zone_key = self.combobox.currentText()

		try:
			with open('data_atr.pickle', 'rb') as file:
				atr_dict = pickle.load(file)
			path_f = atr_dict['path']
		except:
			path_f = os.getcwd()
		if self.total:
			if check_extension('.gpx', self.total):
				path = QFileDialog.getSaveFileName(self, 'Сохранить файл', path_f, 'KML(*.kml)')
				if path[0] != '':
					gpx_to_kml(self.total, path[0])
					msg = QMessageBox.information(self, 'Сообщение', 'GPX файл конвертирован в KML файл', QMessageBox.Ok)

					atr_dict["path"] = os.path.dirname(path[0])

					with open('data_atr.pickle', 'wb') as file:
						pickle.dump(atr_dict, file)

			elif check_extension('.xls', self.total) and len(self.total) == 1:
				path = QFileDialog.getSaveFileName(self, "Сохранить файл", path_f, "KML(*.kml)" )
				if path[0] != '':
					xls_to_kml(self.total, path[0], zone_key)
					msg = QMessageBox.information(self, 'Сообщение', 'Excel файл сконвертирован в KML файл с зоной ' + zone_key, QMessageBox.Ok)

					atr_dict["path"] = os.path.dirname(path[0])

					with open('data_atr.pickle', 'wb') as file:
						pickle.dump(atr_dict, file)
			else:
				QMessageBox.information(self, 'Сообщение', 'Неверный формат файлов', QMessageBox.Ok)
		else:
			QMessageBox.information(self, 'Сообщение', 'Пустая таблица', QMessageBox.Ok)

	def on_btnKmlToGpx(self):
		zone_key = self.combobox.currentText()

		try:
			with open('data_atr.pickle', 'rb') as file:
				atr_dict = pickle.load(file)
			path_f = atr_dict['path']
		except:
			path_f = os.getcwd()
		if self.total:
			if check_extension('.kml', self.total):
				path = QFileDialog.getSaveFileName(self, 'Сохранить файл', path_f, 'GPX(*.gpx)')
				if path[0] != '':
					kml_to_gpx(self.total, path[0])
					msg = QMessageBox.information(self, 'Сообщение', 'KML файл конвертирован в GPX файл', QMessageBox.Ok)

					atr_dict["path"] = os.path.dirname(path[0])

					with open('data_atr.pickle', 'wb') as file:
						pickle.dump(atr_dict, file)

			elif check_extension('.xls', self.total) and len(self.total) == 1:
				path = QFileDialog.getSaveFileName(self, 'Сохранить файл', path_f, 'GPX(*.gpx)')
				if path[0] != '':
					xls_to_gpx(self.total, path[0], zone_key)
					msg = QMessageBox.information(self, "Сообщение", "Excel файл конвертирован в GPX файл", QMessageBox.Ok)

					atr_dict["path"] = os.path.dirname(path[0])

					with open('data_atr.pickle', 'wb') as file:
						pickle.dump(atr_dict, file)
		else:
			QMessageBox.information(self, "Сообщение", 'Пустая таблица', QMessageBox.Ok)

	def on_btnMerge(self):
		try:
			with open('data_atr.pickle', 'rb') as file:
				atr_dict = pickle.load(file)
			path_f = atr_dict["path"]
		except:
			path_f = os.getcwd()

		if self.total:
			if check_extension('.xls', self.total):
				path = QFileDialog.getSaveFileName(self, 'Сохранить файл', path_f, 'XLS(*.xls)')
				if path[0] != '':
					# try:
					# 	total_table(self.total, path[0], self.errParam)
					# except:
					# 	QMessageBox.information(self, 'Сообщение', 'Что-то пошло не так. Проверьте правильность заполнения таблиц', QMessageBox.Ok)
					# else:
					# 	QMessageBox.information(self, 'Сообщение', 'Общая таблица была создана', QMessageBox.Ok)

					# 	atr_dict["path"] = os.path.dirname(path[0])

					# 	with open('data_atr.pickle', 'wb') as file:
					# 		pickle.dump(atr_dict, file)	
					response = total_table(self.total, path[0], self.errParam)

					if response == "success":
						QMessageBox.information(self, 'Сообщение', 'Общая таблица была создана', QMessageBox.Ok)
					else:
						QMessageBox.information(self, "Сообщение" ,"Ошибка! " + response, QMessageBox.Ok)
					
					atr_dict["path"] = os.path.dirname(path[0])
					with open('data_atr.pickle', 'wb') as file:
						pickle.dump(atr_dict, file)
			else:
				QMessageBox.information(self, 'Сообщение', 'Все файлы должны иметь расширение "xls"', QMessageBox.Ok)
		else:
			QMessageBox.information(self, "Сообщение", 'Пустая таблица', QMessageBox.Ok)
		
	# End

	# Table events
	def cell_was_clicked(self, row, column):
		modifiers = QApplication.keyboardModifiers()

		if modifiers == Qt.ShiftModifier:
			if type(self.current_selected_row) == int:
				if column == 0:
					self.table.setRangeSelected(QTableWidgetSelectionRange(row, column + 1, self.current_selected_row, column), True)
				else:
					self.table.setRangeSelected(QTableWidgetSelectionRange(row, column - 1, self.current_selected_row, column), True)
		else:
			if column == 0:
				self.table.setRangeSelected(QTableWidgetSelectionRange(row, column + 1, row, column + 1), True)
			else:
				self.table.setRangeSelected(QTableWidgetSelectionRange(row, column - 1, row, column - 1), True)
			self.current_selected_row = row
	# End

	# Key events
	def keyPressEvent(self, e):
		if e.key() == Qt.Key_Delete:
			self.total = []
			self.table.clearContents()
	# End

	# Functions
	def from_list_to_table(self, list_of_files):
		'''Move data from list of file names to self.table'''
		for ind, file in enumerate(list_of_files):
			rel = file.split('/')[-1]
			self.table.setItem(ind, 0, QTableWidgetItem(rel))
			self.table.item(ind, 0).setTextAlignment(Qt.AlignCenter)
			statbuf = os.path.getmtime(file)
			st = str(datetime.fromtimestamp(statbuf))
			st = st if not '.' in st else st.split('.')[0] 
			self.table.setItem(ind, 1, QTableWidgetItem(st))
			self.table.item(ind, 1).setTextAlignment(Qt.AlignCenter)
	# End


if __name__ == '__main__':
	app = QApplication(sys.argv)
	win = Window()
	win.show()
	sys.exit(app.exec_())




















