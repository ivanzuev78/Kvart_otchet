import json

from Otchet_class import Cunductor
from PyQt5 import uic, QtWidgets, QtGui
from PyQt5.QtWidgets import QTableWidgetItem, QFileDialog
from PyQt5.QtCore import QDate
from copy import copy
from time import localtime, struct_time
import os
import pickle


# Главное окно
class MainWindow(QtWidgets.QMainWindow, uic.loadUiType("main_window.ui")[0]):
    """
    Главное окно
    """
    kvartal_dict = {1: {'start': [1, 1], 'end': [3, 31]},
                    2: {'start': [4, 1], 'end': [6, 30]},
                    3: {'start': [7, 1], 'end': [9, 30]},
                    4: {'start': [10, 1], 'end': [12, 31]}}

    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)

        self.unknown_themes = []
        self.theme_list = {}
        self.known_themes = []
        self.ignore_name_list = []
        self.ignore_tema_list = []
        self.text_of_tema = {}
        self.text_of_tema_noname = {}
        self.read_xl = True

        self.year_numb = None
        self.kvartal_numb = None
        self.window = None
        self.file_settings_window = None
        self.date_window = None
        self.personal_date = False
        self.update_bot.clicked.connect(self.update_it)
        self.make_xl_bot.clicked.connect(self.make_xl)
        self.make_xl_bot_2.clicked.connect(self.make_xl_noname)
        self.settings.clicked.connect(self.settings_window)
        self.read_data_but.clicked.connect(self.read_data)
        self.file_settings_wg.clicked.connect(self.file_settings)
        self.change_date_but.clicked.connect(self.change_date_window)
        self.checkBox_date.clicked.connect(self.date_changer)
        self.change_date_but.setEnabled(False)

        self.show_data.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.path_program = os.getcwd()

        self.set_kvartal()
        self.counductor = Cunductor()
        self.show()
        self.load_settings()

    def debug(self):
        pass

    def date_changer(self):
        if self.checkBox_date.isChecked():
            self.change_date_but.setEnabled(True)
            self.year.setEnabled(False)
            self.kvartal.setEnabled(False)
            self.personal_date = True
        else:
            self.change_date_but.setEnabled(False)
            self.year.setEnabled(True)
            self.kvartal.setEnabled(True)
            self.personal_date = False

    # Вызывает окно настройки файла
    def file_settings(self):
        if not self.file_settings_window:
            self.file_settings_window = ChooseFile(self)
        self.setEnabled(False)
        self.file_settings_window.show()

    # Вызывает окно настройки тем
    def settings_window(self):
        if not self.window:
            self.window = SecondWindow(self)
        self.setEnabled(False)
        self.window.show()

    def change_date_window(self):
        if not self.date_window:
            self.date_window = DateWindow(self)

        self.setEnabled(False)
        self.date_window.show()

    # Кнопка "Считать данные"
    def read_data(self):
        self.current_action.setText('Считывание данных')
        for process, text in self.counductor.reader():
            self.progressBar.setValue(process)
            self.current_action.setText(text)

        self.current_action.setText('Данные считаны')

    # Кнопка "Обновить данные"
    def update_it(self):

        self.update_date()
        self.counductor.counter()
        self.counductor.sort_data()

        self.show_data.setRowCount(len(self.counductor.thems))
        for row, tema in enumerate(self.counductor.thems):
            self.show_data.setItem(row, 0, QTableWidgetItem(f"{tema}"))
            self.show_data.setItem(row, 1, QTableWidgetItem(f"{self.counductor.thems[tema].total_sintes_count} | "
                                                            f"{self.counductor.thems[tema].total_sintes_mass} кг"))
            self.show_data.setItem(row, 2, QTableWidgetItem(f"     {self.counductor.thems[tema].total_obraz}"))
            self.show_data.setItem(row, 3, QTableWidgetItem(f"      {self.counductor.thems[tema].total_otchet}"))
            self.show_data.setItem(row, 4, QTableWidgetItem(f"        {self.counductor.thems[tema].total_nanesenie}"))
            self.progressBar.setValue(100)

        self.show_data.resizeColumnsToContents()

    # Функция считывания обновления дат
    def update_date(self):
        if self.personal_date:
            self.counductor.change_date(self.date_window.start_date_numb, self.date_window.end_date_numb)
            pass
        else:
            self.year_numb = self.year.date().year()
            self.kvartal_numb = self.kvartal.currentIndex() + 1
            self.counductor.change_date([self.year_numb] + self.kvartal_dict[self.kvartal_numb]['start'],
                                        [self.year_numb] + self.kvartal_dict[self.kvartal_numb]['end'])

    # Создает отчёт с именами
    def make_xl(self):
        try:
            path = self.counductor.make_excel(self.kvartal_numb, self.year_numb, self.personal_date)
            self.current_action.setText('Отчёт сформирован')
            if self.open_check.isChecked():
                os.startfile(path)
        except:
            self.current_action.setText('Закройте файл Excel')

    # Создает отчёт без имён
    def make_xl_noname(self):
        try:
            path = self.counductor.make_excel_noname(self.kvartal_numb, self.year_numb, self.personal_date)
            self.current_action.setText('Отчёт сформирован')
            if self.open_check_2.isChecked():
                os.startfile(path)
        except:
            self.current_action.setText('Закройте файл Excel')

    # Загружает настройки
    def load_settings(self):
        if os.path.exists('settings.json'):
            try:
                with open('settings.json', 'r') as f:

                    to_load = json.load(f)

                    self.counductor.ignor_tema = to_load['ignor_tema']
                    self.counductor.ignor_names = to_load['ignor_names']
                    self.counductor.text_of_tema_noname = to_load['text_of_tema_noname']
                    self.counductor.text_of_tema = to_load['text_of_tema']
                    self.counductor.global_tems = to_load['global_tems']
                    self.counductor.svodnaya_name_file = to_load['svodnaya_name_file']
                    self.counductor.production_name_file = to_load['production_name_file']
                    self.counductor.report_name_file = to_load['report_name_file']
                    self.counductor.sintez_name_file = to_load['sintez_name_file']
                    self.counductor.thems_replased = to_load['thems_replased']
                    self.counductor.good_names = to_load['good_names']

            except:
                self.error_dialog = QtWidgets.QErrorMessage()
                self.error_dialog.showMessage('Ошибка загрузки настроек. Установленны настройки по умолчанию')
        else:
            self.error_dialog = QtWidgets.QErrorMessage()
            self.error_dialog.showMessage('Файл с настройками \"settings.json\" не найдет. Установленны настройки по умолчанию')

    # Устанавливает нужный квартал при запуске программы
    def set_kvartal(self):
        if localtime()[1] > 1:
            self.year.setDate(QDate(localtime()[0], 1, 1))
        else:
            self.year.setDate(QDate(localtime()[0] - 1, 1, 1))

        data1 = struct_time((localtime()[0], 2, 1, 1, 1, 1, 1, 1, 1))
        data2 = struct_time((localtime()[0], 5, 1, 1, 1, 1, 1, 1, 1))
        data3 = struct_time((localtime()[0], 8, 1, 1, 1, 1, 1, 1, 1))
        data4 = struct_time((localtime()[0], 11, 1, 1, 1, 1, 1, 1, 1))

        if data1 <= localtime() < data2:
            self.kvartal.setCurrentIndex(0)
        elif data2 <= localtime() < data3:
            self.kvartal.setCurrentIndex(1)
        elif data3 <= localtime() < data4:
            self.kvartal.setCurrentIndex(2)
        else:
            self.kvartal.setCurrentIndex(3)

    # Сохраняет настройки
    def save_settings(self):
        to_save = {'ignor_tema': self.counductor.ignor_tema,
                   'ignor_names': self.counductor.ignor_names,
                   'text_of_tema_noname': self.counductor.text_of_tema_noname,
                   'text_of_tema': self.counductor.text_of_tema,
                   'global_tems': self.counductor.global_tems,
                   'svodnaya_name_file': self.counductor.svodnaya_name_file,
                   'production_name_file': self.counductor.production_name_file,
                   'report_name_file': self.counductor.report_name_file,
                   'sintez_name_file': self.counductor.sintez_name_file,
                   'thems_replased': self.counductor.thems_replased,
                   'good_names': self.counductor.good_names
                   }
        with open('settings.json', 'w') as f:
            json.dump(to_save, f, indent=4, ensure_ascii=False)


# Окно настройки тем
class SecondWindow(QtWidgets.QMainWindow, uic.loadUiType("settings.ui")[0]):
    def __init__(self, main_wind):
        self.main_window = main_wind
        super(SecondWindow, self).__init__()
        self.setupUi(self)
        self.edit_names_window = None

        self.close_but.clicked.connect(self.close)
        self.tema_add_But.clicked.connect(self.add_global_tema)
        self.move_tema_but.clicked.connect(self.move_tema)
        self.remove_tema_but.clicked.connect(self.remove_tema)
        self.del_tema_but.clicked.connect(self.remove_global_tema)
        self.ignore_tema_but.clicked.connect(self.add_ignor_tema)
        self.not_ignor_tema_but.clicked.connect(self.remove_ignor_tema)
        self.official_tems.itemSelectionChanged.connect(self.update_not_official_tems)
        self.save_text_but.clicked.connect(self.save_text)
        self.ignore_name_but.clicked.connect(self.add_ignor_name)
        self.not_ignore_name_but.clicked.connect(self.remove_ignor_name)
        self.edit_names_but.clicked.connect(self.edit_names)

    # Скрипт, запускаемый при показе окна
    def showEvent(self, a0: QtGui.QShowEvent) -> None:
        self.clear_all()
        self.tema_navigator()
        self.update_names()

    # Очищает все поля. Нужна при вызове окна
    def clear_all(self):
        while self.ignore_tems.takeItem(0):
            pass
        while self.unkown_tema.takeItem(0):
            pass
        while self.official_tems.takeItem(0):
            pass
        while self.not_official_tems.takeItem(0):
            pass
        while self.names_widget.takeItem(0):
            pass
        while self.ignore_name_wg.takeItem(0):
            pass

    # Скрипт, запускаемый при закрытии окна
    def closeEvent(self, event):
        self.main_window.setEnabled(True)
        self.main_window.save_settings()
        self.main_window.update_it()
        event.accept()

    # Скрипт, который заполняет все поля при показе окна
    def tema_navigator(self):
        for tema in self.main_window.counductor.thems:
            if tema in self.main_window.counductor.ignor_tema:
                continue

            for thema_replased in self.main_window.counductor.thems_replased:
                if tema in thema_replased:
                    continue

            if tema in self.main_window.counductor.global_tems:
                continue

            self.unkown_tema.addItem(tema)

        for globa_tema in self.main_window.counductor.global_tems:
            self.official_tems.addItem(globa_tema)

        for ignor_tema in self.main_window.counductor.current_ignor_tema:
            self.ignore_tems.addItem(ignor_tema)

    # Заполняет окно с именами
    def update_names(self):
        for name in self.main_window.counductor.names:
            if name in self.main_window.counductor.ignor_names:
                self.ignore_name_wg.addItem(name)
            else:
                self.names_widget.addItem(self.main_window.counductor.input_name(name))

    # Добавляет имя в игнор
    def add_ignor_name(self):
        try:
            ignor_name = self.names_widget.takeItem(self.names_widget.currentRow()).text()
            self.main_window.counductor.ignor_names.append(ignor_name)
            self.ignore_name_wg.addItem(ignor_name)
        except:
            pass

    # Убирает имя из игнора
    def remove_ignor_name(self):
        try:
            ignor_name = self.ignore_name_wg.takeItem(self.ignore_name_wg.currentRow()).text()
            del self.main_window.counductor.ignor_names[self.main_window.counductor.ignor_names.index(ignor_name)]
            self.names_widget.addItem(ignor_name)
        except:
            pass

    # Добавляет тему в игнор
    def add_ignor_tema(self):
        try:
            tema_fly = self.unkown_tema.takeItem(self.unkown_tema.currentRow()).text()

            self.main_window.counductor.ignor_tema.append(tema_fly)
            self.ignore_tems.addItem(tema_fly)
        except:
            pass

    # Удаляет тему из игнора
    def remove_ignor_tema(self):
        try:
            tema_to_del = self.ignore_tems.takeItem(self.ignore_tems.currentRow()).text()
            self.main_window.counductor.ignor_tema.pop(self.main_window.counductor.ignor_tema.index(tema_to_del))
            self.unkown_tema.addItem(tema_to_del)
        except:
            pass

    # Добавляет глобальную тему
    def add_global_tema(self):
        if self.tema_add_line.text() and str(self.tema_add_line.text()) not in self.main_window.counductor.global_tems:
            global_tema = self.tema_add_line.text()
            self.official_tems.addItem(global_tema)
            self.main_window.counductor.global_tems[global_tema] = []

        self.tema_add_line.setText('')

    # Удаляет глобальную тему
    def remove_global_tema(self):
        global_tema = self.official_tems.takeItem(self.official_tems.currentRow())
        if global_tema:
            tema_to_del = global_tema.text()
            for tema in copy(self.main_window.counductor.thems_replased):
                if tema_to_del == self.main_window.counductor.thems_replased[tema]:
                    self.unkown_tema.addItem(tema)
                    del self.main_window.counductor.thems_replased[tema]

                if tema_to_del in self.main_window.counductor.text_of_tema:
                    del self.main_window.counductor.text_of_tema[tema_to_del]

                if tema_to_del in self.main_window.counductor.text_of_tema_noname:
                    del self.main_window.counductor.text_of_tema_noname[tema_to_del]

            del self.main_window.counductor.global_tems[tema_to_del]

    # Перемещает тему в список глобальных тем
    def move_tema(self):
        try:
            global_tema = self.official_tems.selectedItems()[0].text()
            if global_tema:
                tema_fly = self.unkown_tema.takeItem(self.unkown_tema.currentRow()).text()
                self.main_window.counductor.thems_replased[tema_fly] = global_tema
                self.main_window.counductor.global_tems[global_tema].append(tema_fly)
                self.update_not_official_tems()

        except:
            pass

    # Удаляет тему из глобальной темы
    def remove_tema(self):
        try:
            global_tema = self.official_tems.selectedItems()[0].text()
            if self.not_official_tems.selectedItems()[0].text():
                tema_fly = self.not_official_tems.takeItem(self.not_official_tems.currentRow()).text()

                del self.main_window.counductor.thems_replased[tema_fly]

                self.main_window.counductor.global_tems[global_tema].pop(
                    self.main_window.counductor.global_tems[global_tema].index(tema_fly))

                self.unkown_tema.addItem(tema_fly)
        except:
            pass

    # Обновляет список неофициальных тем. Нужно при листании глобальных тем
    def update_not_official_tems(self):
        try:
            while self.not_official_tems.takeItem(0):
                pass

            global_tema = self.official_tems.selectedItems()[0].text()

            if global_tema:
                for tema in self.main_window.counductor.global_tems[global_tema]:
                    # if tema in self.main_window.counductor.all_thems:
                    self.not_official_tems.addItem(tema)

                if global_tema in self.main_window.counductor.text_of_tema:
                    self.wd_text_of_tema.setPlainText(
                        self.main_window.counductor.text_of_tema[global_tema])
                else:
                    self.wd_text_of_tema.setPlainText('')

                if global_tema in self.main_window.counductor.text_of_tema_noname:
                    self.wd_text_of_tema_noname.setPlainText(
                        self.main_window.counductor.text_of_tema_noname[global_tema])
                else:
                    self.wd_text_of_tema_noname.setPlainText('')
        except:
            pass

    # Сохраняет текст к глобальной теме
    def save_text(self):

        if self.official_tems.selectedItems():
            if self.wd_text_of_tema.toPlainText():
                self.main_window.counductor.text_of_tema[self.official_tems.selectedItems()[0].text()] = \
                    self.wd_text_of_tema.toPlainText()
            if self.wd_text_of_tema_noname.toPlainText():
                self.main_window.counductor.text_of_tema_noname[self.official_tems.selectedItems()[0].text()] = \
                    self.wd_text_of_tema_noname.toPlainText()

    def edit_names(self):
        if not self.edit_names_window:
            self.edit_names_window = EditNameWindow(self)
        self.setEnabled(False)
        self.edit_names_window.show()


# Окно настройки пути поиска файлов
class ChooseFile(QtWidgets.QMainWindow, uic.loadUiType("choose_file.ui")[0]):
    def __init__(self, main_wind):
        self.main_window = main_wind
        super(ChooseFile, self).__init__()
        self.setupUi(self)

        # Соединяем кнопки с соответствующими функциями
        self.choose_svod_wg.clicked.connect(self.choose_file('svod'))
        self.choose_prod_wg.clicked.connect(self.choose_file('prod'))
        self.choose_otchet_wg.clicked.connect(self.choose_file('otchet'))
        self.choose_sintez_wg.clicked.connect(self.choose_file('sintez'))

        self.close_but.clicked.connect(self.close)

        # Прописываем текущие местоположения файлов в поля рядом с кнопками
        self.label_svod.setText(f'{self.main_window.counductor.svodnaya_name_file}')
        self.label_prod.setText(f'{self.main_window.counductor.production_name_file}')
        self.label_otchet.setText(f'{self.main_window.counductor.report_name_file}')
        self.label_sintez.setText(f'{self.main_window.counductor.sintez_name_file}')

    # Функиця, вызывающая окно выбора файла
    def choose_file(self, name_file):

        def wraper():
            headlines = {
                'svod': "Выбрать файл сводной таблицы",
                'prod': "Выбрать файл перечня продцкции ОВНТ",
                'otchet': "Выбрать файл перечня отчётов",
                'sintez': "Выбрать файл списка синтезов"
            }
            filename, _ = QFileDialog.getOpenFileName(self, headlines[name_file])
            if filename:
                if name_file == 'svod':
                    self.main_window.counductor.svodnaya_name_file = filename
                    self.label_svod.setText(f'{filename}')
                elif name_file == 'prod':
                    self.main_window.counductor.production_name_file = filename
                    self.label_prod.setText(f'{filename}')
                elif name_file == 'otchet':
                    self.main_window.counductor.report_name_file = filename
                    self.label_otchet.setText(f'{filename}')
                elif name_file == 'sintez':
                    self.main_window.counductor.sintez_name_file = filename
                    self.label_sintez.setText(f'{filename}')

        return wraper

    # Скрипт, срабатывающий при закрытии окна
    def closeEvent(self, event) -> None:
        self.main_window.setEnabled(True)
        self.main_window.save_settings()


# Окно изменения персональных дат
class DateWindow(QtWidgets.QMainWindow, uic.loadUiType("Change_date.ui")[0]):
    def __init__(self, main_wind: MainWindow):
        self.main_window = main_wind
        super(DateWindow, self).__init__()
        self.setupUi(self)
        self.save_but.clicked.connect(self.save)

        self.start_date_numb = [2020, 1, 1]
        self.end_date_numb = [2020, 12, 31]

    def save(self):
        self.start_date_numb = [self.start_date.date().year(), self.start_date.date().month(),
                                self.start_date.date().day()]
        self.end_date_numb = [self.end_date.date().year(), self.end_date.date().month(),
                              self.end_date.date().day()]
        self.main_window.update_date()
        self.close()

        pass

    def closeEvent(self, a0: QtGui.QCloseEvent) -> None:
        self.main_window.setEnabled(True)


# Окно редактирования имён
class EditNameWindow(QtWidgets.QMainWindow, uic.loadUiType("edit_names.ui")[0]):
    def __init__(self, settings_window: MainWindow):
        self.settings_window = settings_window
        super(EditNameWindow, self).__init__()
        self.setupUi(self)
        self.save_but.clicked.connect(self.close)
        self.add_name_but.clicked.connect(self.add_name)
        self.del_name_but.clicked.connect(self.del_name)

    def fill_table(self, names_len=None):
        while self.all_names_widget.takeItem(0,0):
            pass
        if names_len is None:
            self.all_names_widget.setRowCount(len(self.settings_window.main_window.counductor.good_names))
        else:
            self.all_names_widget.setRowCount(names_len)

        for row, name in enumerate(self.settings_window.main_window.counductor.good_names):
            self.all_names_widget.setItem(row, 0, QTableWidgetItem(f"{name}"))
            self.all_names_widget.setItem(row, 1, QTableWidgetItem(f"{self.settings_window.main_window.counductor.good_names[name]}"))

    def add_name(self):
        numb_of_names_now = len(self.settings_window.main_window.counductor.good_names)
        good_name = self.good_name.text()
        bad_name = self.bad_name.text()
        self.fill_table(names_len = numb_of_names_now + 1)
        self.all_names_widget.setItem(numb_of_names_now, 0, QTableWidgetItem(
            f"{bad_name}"))
        self.all_names_widget.setItem(numb_of_names_now, 1, QTableWidgetItem(
            f"{good_name}"))
        self.settings_window.main_window.counductor.good_names[bad_name] = good_name


    def del_name(self):
        try:
            row = self.all_names_widget.currentRow()
            bad_name = self.all_names_widget.item(row, 0).text()
            del self.settings_window.main_window.counductor.good_names[bad_name]
            self.fill_table()
        except:
            pass

    def showEvent(self, a0: QtGui.QShowEvent) -> None:
        self.fill_table()

    def closeEvent(self, a0: QtGui.QCloseEvent) -> None:
        self.settings_window.setEnabled(True)