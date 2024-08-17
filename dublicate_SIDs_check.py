import openpyxl
import subprocess
import time
import funcs
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget, QTextEdit, QDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QThread, pyqtSignal
import os.path

# Создаем переменные для работы с таблицей compare_sids_xlsx
compare_sids_xlsx, sheet = funcs.table_var_assign()

# Получаем массивы с именами и описаниями WS из workstations.txt
ws_list, ws_list_discr = funcs.get_ws_names_and_discr()

n = 0

# Создаем класс диалогового окна кнопки создания новой таблицы compare_sids.xlsx
class CreateTableWindow(QDialog):
    def __init__(self, parent, title, x, y):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedWidth(x)
        self.setFixedHeight(y)
        self.setWindowIcon(QIcon('icon.ico'))
        self.center()

        # Создаем вертикальный layout
        layout = QVBoxLayout()

        # Добавляем текстовое поле
        self.label = QLabel(f'<center>Таблица "compare_sids.xlsx"'
                            f'<p>уже существует. Перезаписать?</center>')
        layout.addWidget(self.label)

        # Создаем горизонтальный layout для кнопок
        button_layout = QHBoxLayout()

        # Добавляем кнопку "Да"
        self.yes_button = QPushButton("Да")
        button_layout.addWidget(self.yes_button)

        # Добавляем кнопку "Нет"
        self.no_button = QPushButton("Нет")
        button_layout.addWidget(self.no_button)

        # Добавляем горизонтальный layout с кнопками в основной вертикальный layout
        layout.addLayout(button_layout)

        # Подключаем сигнал к слоту (обработчику)
        self.yes_button.clicked.connect(self.close)
        self.no_button.clicked.connect(self.close)

        # Устанавливаем layout для окна
        self.setLayout(layout)

    # Метод для центрирования окна относительно родительского
    def center(self):
        parent_geometry = self.parent().geometry()
        self.setGeometry(
            parent_geometry.x() + parent_geometry.width() // 2 - self.width() // 2,
            parent_geometry.y() + parent_geometry.height() // 2 - self.height() // 2,
            self.width(),
            self.height()
        )

# Создаем класс диалогового окна кнопки "About"
class AboutWindow(QDialog):
    def __init__(self, parent, title, x, y):
        super().__init__(parent)
        self.setWindowTitle(title)  # Устанавливаем заголовок окна
        self.setFixedWidth(x)  # Устанавливаем фиксированную ширину окна
        self.setFixedHeight(y)  # Устанавливаем фиксированную высоту окна
        self.setWindowIcon(QIcon('icon.ico'))  # Устанавливаем иконку окна
        self.center()  # Центрируем окно

        # Создаем вертикальный layout
        layout = QVBoxLayout()

        # Добавляем текстовое поле с инструкцией
        self.label = QLabel(f'<p><center><img src="sid.png" width=100 height=100/></center>'
                            f'<p><h3><center>Dublicate SIDs check</h3></center>'
                            f'<p><center>С помощью этой программы можно выявить дублирующиеся'
                            f'<p>Domain и Local SIDs в локальной сети</center>'
                            f'<p>Инструкция по использованию:'
                            f'<p>1) Экспортируем каталог с компьютерами из Active Directory в "workstations.txt".'
                            f'<p>2) Помещаем файл в директорию с программой.'
                            f'<p>3) В той же директории должна находится таблица "compare_sids.xlsx" с определенным шаблоном'
                            f'<p>4) Запускаем от имени администратора.'
                            f'<p>5) Запускаем "Получить SIDs" и ждем завершения.'
                            f'<p>6) Запускаем "Проверка на дублирование SIDs" и ждем завершения.'
                            f'<p>7) Вывод результатов производится в "compare_sids.xlsx"')
        layout.addWidget(self.label)

        # Добавляем кнопку для закрытия окна
        self.close_button = QPushButton("Закрыть")
        layout.addWidget(self.close_button)

        # Подключаем сигнал к слоту для закрытия окна
        self.close_button.clicked.connect(self.close)

        # Устанавливаем layout для окна
        self.setLayout(layout)

    # Метод для центрирования окна относительно родительского
    def center(self):
        parent_geometry = self.parent().geometry()
        self.setGeometry(
            parent_geometry.x() + parent_geometry.width() // 2 - self.width() // 2,
            parent_geometry.y() + parent_geometry.height() // 2 - self.height() // 2,
            self.width(),
            self.height()
        )

# Класс для выполнения длительных операций в отдельном потоке
class WorkerThread(QThread):
    # Сигналы для передачи результатов
    result_ready = pyqtSignal(object)  # Сигнал для полного результата
    partial_result_ready = pyqtSignal(tuple)  # Сигнал для частичных результатов

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func  # Функция для выполнения в потоке
        self.args = args  # Аргументы функции
        self.kwargs = kwargs  # Ключевые аргументы функции

    # Метод выполнения потока
    def run(self):
        results = []  # Список для хранения результатов
        for result in self.func(*self.args, **self.kwargs):
            self.partial_result_ready.emit(result)  # Отправляем частичный результат
            results.append(result)  # Добавляем результат в список
        self.result_ready.emit(results)  # Отправляем полный результат

# Создаем класс основного окна
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dublicate SIDs Check")
        self.setGeometry(100, 100, 400, 500)
        self.setFixedWidth(400)
        self.setFixedHeight(500)
        self.setWindowIcon(QIcon('icon.ico'))

        # Создаем центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Создаем вертикальный layout
        layout = QVBoxLayout()

        # Добавляем кнопку "About"
        self.help_for_use_button = QPushButton("About")
        layout.addWidget(self.help_for_use_button)

        # Добавляем кнопку "Создать "compare_sids.xlsx""
        self.create_table_button = QPushButton('Создать "compare_sids.xlsx"')
        layout.addWidget(self.create_table_button)

        # Добавляем кнопку "Получить SIDs"
        self.get_sids_button = QPushButton("Получить SIDs")
        layout.addWidget(self.get_sids_button)

        # Добавляем кнопку "Получить недостающие SIDs"
        self.get_failed_sids_button = QPushButton("Получить недостающие SIDs")
        layout.addWidget(self.get_failed_sids_button)

        # Добавляем кнопку "Проверка на дублирование SIDs"
        self.get_dubles_button = QPushButton("Проверка на дублирование SIDs")
        layout.addWidget(self.get_dubles_button)

        # Добавляем текстовое поле для вывода логов
        self.message_output = QTextEdit()
        self.message_output.setReadOnly(True)  # Делаем поле только для чтения
        layout.addWidget(self.message_output)

        # Устанавливаем layout для центрального виджета
        central_widget.setLayout(layout)

        # Присваиваем событие нажатия на кнопку к определенной функции и определяем вызов определенного слота
        self.help_for_use_button.clicked.connect(self.help_for_use)
        self.create_table_button.clicked.connect(self.create_compare_sids_table)
        self.get_sids_button.clicked.connect(self.get_domain_n_local_sids)
        # self.get_failed_sids_button.clicked.connect(self.get_failed_domain_n_local_sids)
        self.get_dubles_button.clicked.connect(self.get_dubles)

    def create_compare_sids_table(self):
        # Создаем и открываем окно CreateTableWindow
        self.secondary_window = CreateTableWindow(self, 'Файл уже существует', 250, 150)
        self.secondary_window.exec_()

    def help_for_use(self):
        # Создаем и открываем окно AboutWindow
        self.secondary_window = AboutWindow(self, 'About', 550, 450)
        self.secondary_window.exec_()

    def get_domain_n_local_sids(self):
        # Добавляем сообщение в текстовое поле и делаем кнопку неактивной
        self.message_output.append(f"Идет получение Domain и Local SIDs. Пожалуйста, подождите."
                                   f"<p>************")
        self.get_sids_button.setEnabled(False)
        # Запускаем поток для получения SIDs
        self.sids_thread = WorkerThread(self.run_get_domain_n_local_sids)
        self.sids_thread.partial_result_ready.connect(self.on_partial_sids_received)  # Подключаемся к новому сигналу
        self.sids_thread.result_ready.connect(self.on_sids_received)
        self.sids_thread.start()

    # Функция для выполнения в отдельном потоке
    def run_get_domain_n_local_sids(self):
        # Получаем SIDs для каждой рабочей станции
        for ws_name_output, message, fail_output, local_sid_output, domain_sid_output in funcs.get_names_and_sids(
                ws_list, sheet, compare_sids_xlsx, ws_list_discr):
            yield (ws_name_output, message, fail_output, local_sid_output, domain_sid_output)

    # Метод для обработки частичных результатов
    def on_partial_sids_received(self, result):
        global n
        ws_name_output, message, fail_output, local_sid_output, domain_sid_output = result
        # Добавляем результаты в текстовое поле
        print('TESTING')
        print(ws_list)
        print(n)
        print('TESTING END')
        if message:
            self.message_output.append(ws_name_output + ":" + "<p>Name: " + ws_list[n] + "<p>Local SID - " + message +
                                       '<p>Domain SID - ' + domain_sid_output + "<p>************")
        else:
            self.message_output.append(ws_name_output + ":" + "<p>Name: " + ws_list[n] + "<p>Local SID - " + local_sid_output +
                                       '<p>Domain SID - ' + domain_sid_output + "<p>************")
        n += 1

    # Метод для обработки полного результата
    def on_sids_received(self, results):
        global n
        self.message_output.append('Domain и Local SIDs получены и записаны в "compare_sids.xlsx"')
        self.message_output.append("************")
        n = 0
        self.get_sids_button.setEnabled(True)  # Делаем кнопку активной

    # Метод для получения недостающих Domain и Local SIDs
    # def get_failed_domain_n_local_sids(self):
    #     self.message_output.append(f"Идет получение недостающих Domain и Local SIDs. Пожалуйста, подождите."
    #                                f"<p>************")
    #     self.get_failed_sids_button.setEnabled(False)
    #     self.failed_sids_thread = WorkerThread(self.run_get_failed_domain_n_local_sids)
    #     self.failed_sids_thread.partial_result_ready.connect(self.on_partial_failed_sids_received)  # Подключаемся к новому сигналу
    #     self.failed_sids_thread.result_ready.connect(self.on_failed_sids_received)
    #     self.failed_sids_thread.start()

    # # Функция для выполнения в отдельном потоке
    # def run_get_failed_domain_n_local_sids(self):
    #     for ws_name_output, message, fail_output, local_sid_output, domain_sid_output in funcs.get_names_and_sids(
    #             ws_list, sheet, compare_sids_xlsx, ws_list_discr):
    #         yield (ws_name_output, message, fail_output, local_sid_output, domain_sid_output)
    #
    # # Метод для обработки частичных результатов
    # def on_partial_failed_sids_received(self, result):
    #     ws_name_output, message, fail_output, local_sid_output, domain_sid_output = result
    #     n = len(self.message_output.toPlainText().split("************")) - 1
    #     if message:
    #         self.message_output.append(ws_name_output + ":" + "<p>Name: " + ws_list[n] + "<p>Local SID - " + message +
    #                                    '<p>Domain SID - ' + domain_sid_output + "<p>************")
    #     else:
    #         self.message_output.append(ws_name_output + ":" + "<p>Name: " + ws_list[n] + "<p>Local SID - " + local_sid_output +
    #                                    '<p>Domain SID - ' + domain_sid_output + "<p>************")
    #
    # # Метод для обработки полного результата
    # def on_failed_sids_received(self, results):
    #     self.message_output.append('Domain и Local SIDs получены и записаны в "compare_sids.xlsx"')
    #     self.message_output.append("************")
    #     self.get_failed_sids_button.setEnabled(True)

    # Метод для проверки дублирующихся SIDs
    def get_dubles(self):
        self.message_output.append("Идет поиск дублирующихся SIDs.")  # Добавляем сообщение в текстовое поле
        self.get_dubles_button.setEnabled(False)  # Делаем кнопку неактивной
        # Создаем поток и подключаем обработчики сигналов
        self.dubles_thread = WorkerThread(self.run_get_dubles)
        self.dubles_thread.partial_result_ready.connect(self.on_partial_dubles_received)
        self.dubles_thread.result_ready.connect(self.on_dubles_checked)
        self.dubles_thread.start()

    # Функция для выполнения в отдельном потоке
    def run_get_dubles(self):
        result_from_func = funcs.check_dubles(sheet, compare_sids_xlsx, True)
        result_from_func = funcs.check_dubles(sheet, compare_sids_xlsx, False)

    # Метод для обработки частичных результатов проверки дублей
    def on_partial_dubles_received(self, result_from_func):
        # Пример обработки и вывода промежуточных результатов
        self.message_output.append(result_from_func)  # Здесь можно изменить формат вывода

    # Метод для обработки результата проверки дублирующихся SIDs
    def on_dubles_checked(self, results):
        self.message_output.append('Поиск дублей закончен. Результаты записаны в "compare_sids.xlsx"')
        self.message_output.append("************")
        self.get_dubles_button.setEnabled(True)  # Делаем кнопку активной


app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())
