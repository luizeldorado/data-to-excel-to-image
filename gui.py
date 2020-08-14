# Main file to start GUI

import sys
import os
import pathlib
from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import traceback

from d2e2i import D2E2I


# Get filename relative to script location, so it works when running from other directories and in pyinstaller
def file(name):
	return os.path.join(getattr(sys, '_MEIPASS', os.path.dirname(__file__)), name)


log_filename = file("log.txt")
#log_filename = os.devnull
qtui_main_filename = file("qtui_main.ui")
qttranslation_filename_start = file("translations/qttranslation_")

def main():

	# Prevent problems with pprint in pyw
	if sys.stdout == None:
		sys.stdout = open(log_filename, "w")
	if sys.stderr == None:
		sys.stderr = sys.stdout

	global app
	app = QApplication([])

	# Multithreading
	global thread_pool
	thread_pool = QThreadPool()
	print("maxThreadCount = {}".format(thread_pool.maxThreadCount()))
	
	# Translation
	global current_qtranslator
	current_qtranslator = QTranslator(app)
	app.installTranslator(current_qtranslator)

	# GUI
	form, qtbase = uic.loadUiType(qtui_main_filename)

	# GUIMainWindow must come first so tr is overriden
	dynamic_type = type("GUIDynamicType", (GUIMainWindow, form, qtbase), dict())

	global gui
	gui = dynamic_type()
	gui.setupUi(gui)
	gui.init()

	# Languages
	# TODO: read translations folder for all possible languages
	global locales
	locales = (
		("English", "C"),
		("Português", "pt_BR"),
	)

	# Load language
	global system_locale
	system_locale = QLocale()

	# Store system locale information
	global system_locale_index

	system_locale_index = -1
	for i, (name, code) in enumerate(locales):
		if code == system_locale.name():
			system_locale_index = i
			break

	# Add locales to combo
	# (must block signals because index changes to 0 when adding the first item)
	gui.languageCombo.blockSignals(True)
	for name, code in locales:
		gui.languageCombo.addItem(name, code)
	gui.languageCombo.blockSignals(False)

	# Set selected item in language combo to current language
	# (this will automatically call change_locale)
	current_locale_index = system_locale_index if system_locale_index != -1 else 0
	gui.languageCombo.setCurrentIndex(current_locale_index)

	# Show
	gui.show()
	app.exec()


class Tr():
	def tr(self, string, context, disambiguation=None, n=-1):
		return app.translate(context.__name__, string, disambiguation, n)


class GUIMainWindow(Tr):
	
	def init(self):
		self.d2e2i = D2E2I()
		self.generate_worker = GenerateWorker(self.d2e2i)

		self.generate_worker.signals.done[bool].connect(self.on_generate_done)
		self.generate_worker.signals.done[bool, str].connect(self.on_generate_done)

		self.generate_worker.signals.opening_files.connect(self.on_generate_opening_files)
		self.generate_worker.signals.reading_template_tags.connect(self.on_generate_reading_template_tags)
		self.generate_worker.signals.generating_row.connect(self.on_generate_generating_row)

	def show_message(self, text):
		QMessageBox.information(self, self.tr("D2E2I", __class__), text, QMessageBox.Ok)
	def show_error(self, text):
		QMessageBox.critical(self, self.tr("D2E2I", __class__), text, QMessageBox.Ok)

	def change_locale(self, locale):
		has_loaded = current_qtranslator.load(locale, qttranslation_filename_start)
		print("change_locale: name = {}, has_loaded = {}".format(locale.name(), has_loaded))

		if not has_loaded and locale.name() != "C":
			self.show_error(self.tr("Error! Could not open translation for {0}", __class__).format(locale.name()))
			return

		self.retranslateUi(self)

		if system_locale_index != -1:
			system_locale_text = locales[system_locale_index][0] + self.tr(" (system default)", __class__)
			self.languageCombo.setItemText(system_locale_index, system_locale_text)

	@pyqtSlot(bool)
	@pyqtSlot(bool, str)
	def on_generate_done(self, success, error_reason=''):

		self.generateButton.setEnabled(True)
		self.generateButton.setText(self.tr("Generate", __class__))

		self.generateLabel.setText("")

		if success:
			self.show_message(self.tr("Files generated!", __class__))
		else:
			self.show_error(self.tr("Error! {0}", __class__).format(error_reason))

	def on_generate_opening_files(self):
		self.generateLabel.setText(self.tr("Opening files...", __class__))

	def on_generate_reading_template_tags(self):
		self.generateLabel.setText(self.tr("Reading template tags...", __class__))

	def on_generate_generating_row(self, current, total):
		self.generateLabel.setText(self.tr("Generating row {0} of {1}...", __class__).format(current, total))

	@pyqtSlot()
	def on_templateFileSelectButton_clicked(self):
		file_path, file_type = QFileDialog.getOpenFileName(self,
			self.tr("Open template file", __class__),
			filter=self.tr("Excel Workbook (*.xlsx)", __class__)
		)
		if file_path:
			self.templateFileSelectEdit.setText(file_path)

	@pyqtSlot()
	def on_dataFileSelectButton_clicked(self):
		file_path, file_type = QFileDialog.getOpenFileName(self,
			self.tr("Open data file", __class__),
			filter=self.tr("Excel Workbook (*.xlsx)", __class__)
		)
		if file_path:
			self.dataFileSelectEdit.setText(file_path)

	@pyqtSlot()
	def on_outputFolderSelectButton_clicked(self):
		dir_path = QFileDialog.getExistingDirectory(self,
			self.tr("Select output folder", __class__)
		)
		if dir_path:
			self.outputFolderSelectEdit.setText(dir_path)

	@pyqtSlot()
	def on_generateButton_clicked(self):
		
		self.generateButton.setEnabled(False)
		self.generateButton.setText(self.tr("Generating...", __class__))

		self.d2e2i.file_template = self.templateFileSelectEdit.text()
		self.d2e2i.file_data = self.dataFileSelectEdit.text()
		self.d2e2i.generate_excel_files = self.generateExcelCheck.isChecked()
		self.d2e2i.generate_image_files = self.generateImageCheck.isChecked()
		self.d2e2i.folder_excel = self.outputFolderSelectEdit.text()
		self.d2e2i.folder_image = self.outputFolderSelectEdit.text()
		# self.d2e2i.prefix = "results "

		thread_pool.start(self.generate_worker)

	@pyqtSlot(int)
	def on_languageCombo_currentIndexChanged(self, index):
		language = self.languageCombo.currentData()
		self.change_locale(QLocale(language))

class GenerateWorkerSignals(QObject):
	done = pyqtSignal([bool], [bool, str])
	opening_files = pyqtSignal()
	reading_template_tags = pyqtSignal()
	generating_row = pyqtSignal(int, int)

class GenerateWorker(QRunnable):
	def __init__(self, d2e2i):
		super().__init__()
		self.setAutoDelete(False)
		self.d2e2i = d2e2i
		self.signals = GenerateWorkerSignals()

	@pyqtSlot()
	def run(self):
		
		try:

			self.signals.opening_files.emit()

			if not self.d2e2i.open_files():
				self.signals.done[bool, str].emit(False, self.tr("Error! Template and/or data files don't exist", __class__))
				return

			self.signals.reading_template_tags.emit()
			tag_cells = self.d2e2i.get_tag_cells_in_template()

			pathlib.Path(self.d2e2i.folder_excel).mkdir(parents=True, exist_ok=True)

			for i, data_row in enumerate(self.d2e2i.iter_data_rows()):

				self.signals.generating_row.emit(i, self.d2e2i.number_of_data_rows())

				self.d2e2i.generate_row(data_row, tag_cells)
				
			self.d2e2i.close_files()

			self.signals.done[bool].emit(True)

		except Exception as e:
			self.signals.done[bool, str].emit(False, str(e))

if __name__ == "__main__":
	main()