# main.py
import sys, os
from PyQt5 import QtWidgets, QtCore, QtGui
import pandas as pd
from pathlib import Path
import analysis

APP_NAME = "Статистика"
AUTHOR = "Розробник: Чаплоуцький А.М., кафедра плодівництва і виноградарства УНУ"

class TableModel(QtCore.QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return len(self._data.index)

    def columnCount(self, parent=None):
        return len(self._data.columns)

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if index.isValid():
            if role == QtCore.Qt.DisplayRole:
                return str(self._data.iat[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role):
        if role == QtCore.Qt.DisplayRole:
            if orientation == QtCore.Qt.Horizontal:
                return str(self._data.columns[section])
            else:
                return str(self._data.index[section])

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        # set window icon if present
        ic = Path(__file__).parent / "statystyka_icon_512.png"
        if ic.exists():
            self.setWindowIcon(QtGui.QIcon(str(ic)))
        self.resize(900,600)
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        layout = QtWidgets.QVBoxLayout(central)

        # Buttons
        btn_layout = QtWidgets.QHBoxLayout()
        self.load_btn = QtWidgets.QPushButton("Завантажити CSV")
        self.load_btn.clicked.connect(self.load_csv)
        btn_layout.addWidget(self.load_btn)
        self.run_btn = QtWidgets.QPushButton("Запустити аналіз")
        self.run_btn.clicked.connect(self.run_analysis)
        btn_layout.addWidget(self.run_btn)
        self.report_btn = QtWidgets.QPushButton("Експорт у Word")
        self.report_btn.clicked.connect(self.export_report)
        btn_layout.addWidget(self.report_btn)
        self.about_btn = QtWidgets.QPushButton("Про програму")
        self.about_btn.clicked.connect(self.show_about)
        btn_layout.addWidget(self.about_btn)
        layout.addLayout(btn_layout)

        # Table view
        self.table = QtWidgets.QTableView()
        layout.addWidget(self.table)

        # Log / output
        self.log = QtWidgets.QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        # state
        self.df = pd.DataFrame()

    def load_csv(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Відкрити CSV", ".", "CSV файли (*.csv);;Всі файли (*)")
        if path:
            try:
                self.df = pd.read_csv(path)
                model = TableModel(self.df)
                self.table.setModel(model)
                self.log.append(f"Завантажено {path} ({len(self.df)} рядків)")
            except Exception as e:
                self.log.append("Помилка при завантаженні CSV: " + str(e))

    def run_analysis(self):
        if self.df.empty:
            self.log.append("Дані не завантажені.")
            return
        # basic auto-detect: find numeric column for Value and first three categorical columns as factors
        num_cols = self.df.select_dtypes(include=['number']).columns.tolist()
        cat_cols = self.df.select_dtypes(include=['object','category']).columns.tolist()
        if not num_cols:
            self.log.append("Не знайдено числового стовпця для аналізу.")
            return
        value = num_cols[0]
        try:
            if len(cat_cols) >= 3:
                a = cat_cols[0]; b = cat_cols[1]; c = cat_cols[2]
                model, anova = analysis.three_way_anova(self.df, a, b, c, value, interaction=True)
                title = "Трифакторний ANOVA (з взаємодіями)"
            elif len(cat_cols) == 2:
                a = cat_cols[0]; b = cat_cols[1]
                model, anova = analysis.two_way_anova(self.df, a, b, value, interaction=True)
                title = "Двофакторний ANOVA (з взаємодією)"
            else:
                a = cat_cols[0]
                model, anova = analysis.one_way_anova(self.df, a, value)
                title = "Однофакторний ANOVA"
            self.log.append(f"{title} — Результати:\n{anova}")
            # run shapiro on residuals
            w,p = analysis.shapiro_test(model.resid)
            self.log.append(f"Shapiro–Wilk (залишки): W={w:.4f}, p={p:.4f}")
            # levene
            stat, lp = analysis.levene_test(self.df, a, value)
            self.log.append(f"Levene для {a}: stat={stat:.4f}, p={lp:.4f}")
            # optional post-hoc and plots quick
            try:
                if len(cat_cols) >= 1:
                    tuk = analysis.tukey_hsd(self.df, cat_cols[0], value)
                    self.log.append("Tukey HSD (приклад по першому фактору):\n" + str(tuk.summary()))
            except Exception as e:
                self.log.append("Post-hoc не вдалося: " + str(e))
        except Exception as e:
            self.log.append("Помилка аналізу: " + str(e))

    def export_report(self):
        if self.df.empty:
            self.log.append("Немає даних для звіту.")
            return
        out_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Зберегти звіт Word", "Статистика_звіт.docx", "Файли Word (*.docx)")
        if out_path:
            analyses = {}
            # minimal sample: one-way or two-way as above
            num_cols = self.df.select_dtypes(include=['number']).columns.tolist()
            cat_cols = self.df.select_dtypes(include=['object','category']).columns.tolist()
            value = num_cols[0]
            try:
                if len(cat_cols) >= 3:
                    a=cat_cols[0]; b=cat_cols[1]; c=cat_cols[2]
                    model, anova = analysis.three_way_anova(self.df, a, b, c, value, interaction=True)
                    analyses['Трифакторний ANOVA'] = anova.to_string()
                    resid = model.resid
                elif len(cat_cols) >= 2:
                    a=cat_cols[0]; b=cat_cols[1]
                    model, anova = analysis.two_way_anova(self.df, a, b, value, interaction=True)
                    analyses['Двофакторний ANOVA'] = anova.to_string()
                    resid = model.resid
                else:
                    a=cat_cols[0];
                    model, anova = analysis.one_way_anova(self.df, a, value)
                    analyses['Однофакторний ANOVA'] = anova.to_string()
                    resid = model.resid
                # plots
                tmpdir = Path.cwd() / "temp_plots"
                tmpdir.mkdir(exist_ok=True)
                box = str(tmpdir / "box.png")
                hist = str(tmpdir / "hist.png")
                interaction = str(tmpdir / "interaction.png")
                analysis.plot_box(self.df, a, value, box)
                analysis.plot_hist(resid, hist)
                plots = [box, hist]
                if len(cat_cols) >= 2:
                    analysis.plot_interaction(self.df, a, b, value, interaction)
                    plots.append(interaction)
                analysis.generate_report(out_path, self.df, analyses, plots)
                self.log.append("Звіт збережено: " + out_path)
            except Exception as e:
                self.log.append("Помилка при генерації звіту: " + str(e))

    def show_about(self):
        QtWidgets.QMessageBox.information(self, "Про програму", APP_NAME + "\n\n" + AUTHOR)

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    # set application icon if present
    ic = Path(__file__).parent / "statystyka_icon_512.png"
    if ic.exists():
        app.setWindowIcon(QtGui.QIcon(str(ic)))
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
