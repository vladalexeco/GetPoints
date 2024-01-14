from PyQt5.QtWidgets import (QVBoxLayout, QHBoxLayout, QPushButton, QCheckBox, QDialog, QRadioButton, QSpinBox, QLabel, QSizePolicy)
import os
import pickle

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