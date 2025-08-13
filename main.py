import sys
import os
import pandas as pd
from PyQt5 import QtWidgets, QtGui, QtCore
from analysis import check_normality, one_way_anova, two_way_anova, three_way_anova, correlation, regression, effect_size_eta_squared
from charts import plot_bar, plot_box, plot_pie
from export_word import export_to_word
import tempfile
import matplotlib.pyplot as plt

class SADApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD - Статистичний Аналіз Даних")
        self.setGeometry(200, 200, 900, 600)
        self.setWindowIcon(QtGui.QIcon("icon.ico"))

        self.data_table = QtWidgets.QTableWidget(self)
        self.data_table.setGeometry(20, 50, 600, 400)
        self.data_table.setRowCount(5)
        self.data_table.setColumnCount(5)
        self.data_table.setHorizontalHeaderLabels(["Колонка 1", "Колонка 2", "Колонка 3", "Колонка 4", "Колонка 5"])

        # Кнопки
        btn_add_row = QtWidgets.QPushButton("Додати рядок", self)
        btn_add_row.setGeometry(650, 50, 200, 30)
        btn_add_row.clicked.connect(self.add_row)

        btn_add_col = QtWidgets.QPushButton("Додати стовпець", self)
        btn_add_col.setGeometry(650, 90, 200, 30)
        btn_add_col.clicked.connect(self.add_col)

        btn_analysis = QtWidgets.QPushButton("Аналіз даних", self)
        btn_analysis.setGeometry(650, 130, 200, 30)
        btn_analysis.clicked.connect(self.run_analysis)

        btn_export = QtWidgets.QPushButton("Експорт у Word", self)
        btn_export.setGeometry(650, 170, 200, 30)
        btn_export.clicked.connect(self.export_results)

        btn_about = QtWidgets.QPushButton("Про розробника", self)
        btn_about.setGeometry(780, 10, 100, 30)
        btn_about.clicked.connect(self.show_about)

        self.results = None
        self.df = None

    def add_row(self):
        self.data_table.insertRow(self.data_table.rowCount())

    def add_col(self):
        col_count = self.data_table.columnCount()
        self.data_table.insertColumn(col_count)
        self.data_table.setHorizontalHeaderItem(col_count, QtWidgets.QTableWidgetItem(f"Колонка {col_count+1}"))

    def table_to_df(self):
        rows = self.data_table.rowCount()
        cols = self.data_table.columnCount()
        data = []
        headers = [self.data_table.horizontalHeaderItem(i).text() if self.data_table.horizontalHeaderItem(i) else f"Col{i+1}" for i in range(cols)]
        for row in range(rows):
            row_data = []
            for col in range(cols):
                item = self.data_table.item(row, col)
                if item and item.text() != "":
                    try:
                        row_data.append(float(item.text()))
                    except ValueError:
                        row_data.append(item.text())
                else:
                    row_data.append(None)
            data.append(row_data)
        df = pd.DataFrame(data, columns=headers)
        return df.dropna(how="all")

    def run_analysis(self):
        self.df = self.table_to_df()

        # Перевірка нормальності першої числової колонки
        num_cols = self.df.select_dtypes(include=['float64', 'int64']).columns
        if len(num_cols) == 0:
            QtWidgets.QMessageBox.warning(self, "Помилка", "Немає числових даних для аналізу.")
            return

        p_value = check_normality(self.df[num_cols[0]].dropna())
        if p_value < 0.05:
            QtWidgets.QMessageBox.information(self, "Ненормальний розподіл",
                "Дані не відповідають нормальному розподілу (p<0.05).\n"
                "Рекомендується використати непараметричні методи або трансформувати дані (наприклад, Box-Cox).")
            return

        # Вибір аналізу
        analysis_type, ok = QtWidgets.QInputDialog.getItem(self, "Вибір аналізу",
                                                           "Оберіть метод:",
                                                           ["Однофакторний ANOVA", "Двофакторний ANOVA", "Трифакторний ANOVA",
                                                            "Кореляція", "Регресія", "Сила впливу факторів"],
                                                           0, False)
        if ok:
            if analysis_type == "Однофакторний ANOVA":
                factor, ok1 = QtWidgets.QInputDialog.getText(self, "Фактор", "Введіть назву фактору:")
                value, ok2 = QtWidgets.QInputDialog.getText(self, "Показник", "Введіть назву показника:")
                if ok1 and ok2:
                    self.results = one_way_anova(self.df, factor, value)

            elif analysis_type == "Двофакторний ANOVA":
                f1, _ = QtWidgets.QInputDialog.getText(self, "Фактор 1", "Назва:")
                f2, _ = QtWidgets.QInputDialog.getText(self, "Фактор 2", "Назва:")
                val, _ = QtWidgets.QInputDialog.getText(self, "Показник", "Назва:")
                self.results = two_way_anova(self.df, f1, f2, val)

            elif analysis_type == "Трифакторний ANOVA":
                f1, _ = QtWidgets.QInputDialog.getText(self, "Фактор 1", "Назва:")
                f2, _ = QtWidgets.QInputDialog.getText(self, "Фактор 2", "Назва:")
                f3, _ = QtWidgets.QInputDialog.getText(self, "Фактор 3", "Назва:")
                val, _ = QtWidgets.QInputDialog.getText(self, "Показник", "Назва:")
                self.results = three_way_anova(self.df, f1, f2, f3, val)

            elif analysis_type == "Кореляція":
                c1, _ = QtWidgets.QInputDialog.getText(self, "Колонка 1", "Назва:")
                c2, _ = QtWidgets.QInputDialog.getText(self, "Колонка 2", "Назва:")
                self.results = correlation(self.df, c1, c2)

            elif analysis_type == "Регресія":
                x, _ = QtWidgets.QInputDialog.getText(self, "X", "Назва змінної X:")
                y, _ = QtWidgets.QInputDialog.getText(self, "Y", "Назва змінної Y:")
                self.results = regression(self.df, x, y)

            elif analysis_type == "Сила впливу факторів":
                factor, _ = QtWidgets.QInputDialog.getText(self, "Фактор", "Назва:")
                value, _ = QtWidgets.QInputDialog.getText(self, "Показник", "Назва:")
                self.results = effect_size_eta_squared(self.df, factor, value)

            QtWidgets.QMessageBox.information(self, "Результати", str(self.results))

    def export_results(self):
        if self.df is None or self.results is None:
            QtWidgets.QMessageBox.warning(self, "Помилка", "Немає даних для експорту.")
            return

        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Зберегти у Word", "", "Word files (*.docx)")
        if file_path:
            fig_path = os.path.join(tempfile.gettempdir(), "chart.png")
            plt.figure()
            self.df.select_dtypes(include=['float64', 'int64']).plot()
            plt.savefig(fig_path)
            export_to_word(file_path, "Звіт SAD", self.df, self.results, [fig_path])
            QtWidgets.QMessageBox.information(self, "Успіх", "Експорт завершено.")

    def show_about(self):
        QtWidgets.QMessageBox.information(self, "Про розробника",
            "Розробник: Чаплоуцький А.М.\nКафедра плодівництва і виноградарства УНУ")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = SADApp()
    window.show()
    sys.exit(app.exec_())
