import sys
import os
import pandas as pd
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import (
    QMainWindow, QApplication, QTableWidgetItem,
    QFileDialog, QMessageBox, QInputDialog
)
from analysis import (
    check_normality, one_way_anova, two_way_anova, three_way_anova,
    correlation, regression, effect_size_eta_squared
)
from charts import plot_bar, plot_box, plot_pie
from export_word import export_to_word
import matplotlib.pyplot as plt

class SADApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD — Статистичний аналіз даних")
        self.setGeometry(200, 100, 900, 600)
        self.setWindowIcon(QtGui.QIcon("icon.ico"))

        self.table = QtWidgets.QTableWidget(self)
        self.table.setGeometry(20, 50, 600, 400)
        self.table.setColumnCount(3)
        self.table.setRowCount(5)
        self.table.setHorizontalHeaderLabels(["Фактор A", "Фактор B", "Значення"])

        # Кнопки
        self.btn_add_row = QtWidgets.QPushButton("Додати рядок", self)
        self.btn_add_row.setGeometry(650, 50, 200, 30)
        self.btn_add_row.clicked.connect(self.add_row)

        self.btn_add_col = QtWidgets.QPushButton("Додати стовпець", self)
        self.btn_add_col.setGeometry(650, 90, 200, 30)
        self.btn_add_col.clicked.connect(self.add_column)

        self.btn_analysis = QtWidgets.QPushButton("Аналіз даних", self)
        self.btn_analysis.setGeometry(650, 130, 200, 30)
        self.btn_analysis.clicked.connect(self.run_analysis)

        self.btn_export = QtWidgets.QPushButton("Експорт у Word", self)
        self.btn_export.setGeometry(650, 170, 200, 30)
        self.btn_export.clicked.connect(self.export_results)

        self.btn_about = QtWidgets.QPushButton("Про розробника", self)
        self.btn_about.setGeometry(750, 10, 120, 25)
        self.btn_about.clicked.connect(self.show_about)

        self.results = None
        self.figures = []
        self.df = None

    def add_row(self):
        self.table.insertRow(self.table.rowCount())

    def add_column(self):
        col_count = self.table.columnCount()
        self.table.insertColumn(col_count)
        self.table.setHorizontalHeaderItem(col_count, QTableWidgetItem(f"Колонка {col_count+1}"))

    def get_table_data(self):
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        data = []
        for r in range(rows):
            row_data = []
            for c in range(cols):
                item = self.table.item(r, c)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        df = pd.DataFrame(data, columns=[self.table.horizontalHeaderItem(c).text() for c in range(cols)])
        return df

    def run_analysis(self):
        df = self.get_table_data()

        try:
            numeric_col = df.select_dtypes(include='number').columns[0]
        except:
            numeric_col = df.columns[-1]

        # Автоматична перевірка нормальності
        try:
            p_value = check_normality(df[numeric_col].astype(float))
        except:
            QMessageBox.warning(self, "Помилка", "Не вдалося перевірити нормальність. Перевірте введені дані.")
            return

        if p_value < 0.05:
            QMessageBox.information(
                self, "Дані не нормальні",
                "Результат перевірки Шапіро–Вілка показує, що дані не розподілені нормально (p < 0.05).\n"
                "Рекомендації:\n- Використовуйте непараметричні методи\n- Виконайте перетворення даних (лог, sqrt)"
            )
            return

        # Вибір аналізу
        analysis_type, ok = QInputDialog.getItem(
            self, "Вибір аналізу",
            "Оберіть тип аналізу:",
            ["Однофакторний ANOVA", "Двофакторний ANOVA", "Трифакторний ANOVA",
             "Кореляція", "Регресія", "Сила впливу факторів"],
            0, False
        )
        if not ok:
            return

        if analysis_type == "Однофакторний ANOVA":
            self.results = one_way_anova(df, df.columns[0], numeric_col)
        elif analysis_type == "Двофакторний ANOVA":
            self.results = two_way_anova(df, df.columns[0], df.columns[1], numeric_col)
        elif analysis_type == "Трифакторний ANOVA":
            self.results = three_way_anova(df, df.columns[0], df.columns[1], df.columns[2], numeric_col)
        elif analysis_type == "Кореляція":
            self.results = correlation(df, df.columns[0], df.columns[1])
        elif analysis_type == "Регресія":
            self.results = regression(df, df.columns[0], numeric_col)
        elif analysis_type == "Сила впливу факторів":
            self.results = effect_size_eta_squared(df, [df.columns[0]], numeric_col)

        # Вибір графіка
        chart_type, ok = QInputDialog.getItem(
            self, "Вибір графіка",
            "Оберіть тип графіка:",
            ["Стовпчикова", "Box plot", "Кругова"],
            0, False
        )
        if not ok:
            return

        fig_path = "chart.png"
        if chart_type == "Стовпчикова":
            plot_bar(df, df.columns[0], numeric_col).savefig(fig_path)
        elif chart_type == "Box plot":
            plot_box(df, df.columns[0], numeric_col).savefig(fig_path)
        elif chart_type == "Кругова":
            plot_pie(df, df[df.columns[0]], df[numeric_col].astype(float)).savefig(fig_path)

        self.figures = [fig_path]
        QMessageBox.information(self, "Готово", "Аналіз виконано та графік побудовано.")

    def export_results(self):
        if self.results is None:
            QMessageBox.warning(self, "Помилка", "Спочатку виконайте аналіз.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти звіт", "", "Word Documents (*.docx)")
        if path:
            export_to_word(path, "Результати статистичного аналізу", self.get_table_data(), self.results, self.figures)
            QMessageBox.information(self, "Готово", "Звіт збережено у Word.")

    def show_about(self):
        QMessageBox.information(
            self, "Про розробника",
            "Програму створив Чаплоуцький А.М.\nКафедра плодівництва і виноградарства УНУ"
        )

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SADApp()
    window.show()
    sys.exit(app.exec_())
