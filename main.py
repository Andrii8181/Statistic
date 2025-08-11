# main.py
import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QTableWidgetItem
from PyQt5 import uic
from analysis import run_analysis

class StatystykaApp(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi("main.ui", self)

        self.data = None
        self.loadButton.clicked.connect(self.load_file)
        self.runButton.clicked.connect(self.run_analysis_and_show)

    def load_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Відкрити файл", "", "CSV Files (*.csv)")
        if file_name:
            self.data = pd.read_csv(file_name)
            self.show_data_in_table()

    def show_data_in_table(self):
        if self.data is not None:
            self.tableWidget.setRowCount(self.data.shape[0])
            self.tableWidget.setColumnCount(self.data.shape[1])
            self.tableWidget.setHorizontalHeaderLabels(self.data.columns)
            for i in range(self.data.shape[0]):
                for j in range(self.data.shape[1]):
                    self.tableWidget.setItem(i, j, QTableWidgetItem(str(self.data.iat[i, j])))

    def run_analysis_and_show(self):
        if self.data is None:
            QMessageBox.warning(self, "Помилка", "Будь ласка, спочатку завантажте дані")
            return
        result_path = run_analysis(self.data)
        QMessageBox.information(self, "Готово", f"Результати збережено у файл:\n{result_path}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StatystykaApp()
    window.setWindowTitle("Статистика — Чаплоуцький А.М., кафедра плодівництва і виноградарства УНУ")
    window.show()
    sys.exit(app.exec_())
