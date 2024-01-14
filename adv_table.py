from PyQt5.QtWidgets import QTableWidget, QMessageBox
import os

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