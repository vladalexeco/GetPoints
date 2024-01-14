from functions import *
import pickle
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QApplication, QDesktopWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget,
 QTableWidgetItem, QHeaderView, QFileDialog, QTableWidgetSelectionRange, QMessageBox, QComboBox)
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from dialog import Dialog
from adv_table import AdvTable


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

		self.btnJournal = QPushButton("В журнал")
		self.mskToKml = QPushButton("МСК в Excel")
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
		self.btnJournal.setToolTip("Конвертирует журнал гамма-съемки в формате <b>xls</b> в <b>gpx</b> формат")
		self.mskToKml.setToolTip("Конвертирует файл <b>xls</b> с координатами <b>МСК</b> в <b>gpx</b> формат")
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

		self.hbox_1.addWidget(self.btnJournal)
		
		self.hbox_1.addWidget(self.btnAdd)
		self.hbox_1.addWidget(self.btnDel)
		self.hbox_1.addWidget(self.btnGpxToKml)

		self.hbox_2.addWidget(self.btnParam)
		self.hbox_2.addWidget(self.mskToKml)
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

		self.btnJournal.clicked.connect(self.on_btnJournal)
		self.mskToKml.clicked.connect(self.on_btnMskToKml)
		
		self.btnParam.clicked.connect(self.dialog.exec)

		self.table.cellClicked.connect(self.cell_was_clicked)

	# Button clicked events

	def on_btnMskToKml(self):
		zone_key = self.combobox.currentText()

		try:
			with open('data_atr.pickle', 'rb') as file:
				atr_dict = pickle.load(file)
			path_f = atr_dict['path']
		except:
			path_f = os.getcwd()

		if self.total:
			if check_extension('.xls', self.total) and len(self.total) == 1:
				path = QFileDialog.getSaveFileName(self, "Сохранить файл", path_f, "KML(*.kml)" )
				
				if path[0] != '':
					xls_to_kml(self.total, path[0], zone_key)
					msg = QMessageBox.information(self, 'Сообщение', 'Excel файл сконвертирован в KML файл с зоной ' + zone_key, QMessageBox.Ok)

					atr_dict["path"] = os.path.dirname(path[0])

					with open('data_atr.pickle', 'wb') as file:
						pickle.dump(atr_dict, file)
			else:
				QMessageBox.information(self, 'Сообщение', 'Неверный формат файлов или файлов больше чем один', QMessageBox.Ok)
		else:
			QMessageBox.information(self, 'Сообщение', 'Пустая таблица', QMessageBox.Ok)


	def on_btnJournal(self):
		
		try:
			with open('data_atr.pickle', 'rb') as file:
				atr_dict = pickle.load(file)

			path_f = atr_dict['path']
		except:
			path_f = os.getcwd()


		if self.total: 
			
			if check_extension(".xls", self.total) and len(self.total) == 1:
				
				path = QFileDialog.getSaveFileName(self, 'Сохранить файл', path_f, 'GPX(*.gpx)')
				
				if path[0] != "":
					
					try:
						xls_journal_to_gpx(self.total, path[0])
					except:
						QMessageBox.information(self, "Сообщение", "Что-то пошло не так. Проверьте формат правильность заполнения .xls файла.")
					else:
						QMessageBox.information(self, 'Сообщение', 'Информация сконвертирована в файл .gpx формат', QMessageBox.Ok)

						atr_dict["path"] = os.path.dirname(path[0])

						with open('data_atr.pickle', 'wb') as file:
							pickle.dump(atr_dict, file)
					
			else:
				QMessageBox.information(self, 'Сообщение', 'Неверный формат файла или файлов больше, чем один', QMessageBox.Ok)

		else:
			QMessageBox.information(self, 'Сообщение', 'Пустая таблица', QMessageBox.Ok)


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