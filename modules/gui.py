from PyQt5.QtWidgets import QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget, QTextEdit, QDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QThread, pyqtSignal

from modules import funcs


compare_sids_xlsx, sheet = funcs.table_var_assign()
ws_list, ws_list_discr = funcs.get_ws_names_and_discr()
n = 0


# class for long processes in different thread
class WorkerThread(QThread):
    result_ready = pyqtSignal(object)
    partial_result_ready = pyqtSignal(tuple)

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        results = []
        for result in self.func(*self.args, **self.kwargs):
            self.partial_result_ready.emit(result)
            results.append(result)
        self.result_ready.emit(results)


class AboutWindow(QDialog):
    def __init__(self, parent, title, x, y):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedWidth(x)
        self.setFixedHeight(y)
        self.setWindowIcon(QIcon('icon.ico'))
        self.center()

        layout = QVBoxLayout()

        self.label = QLabel(f'<p><center><img src="sid.png" width=100 height=100/></center>'
                            f'<p><h3><center>Dublicate SIDs check</h3></center>'
                            f'<p><center>С помощью этой программы можно выявить дублирующиеся'
                            f'<p>Domain и Local SIDs в локальной сети</center>'
                            f'<p>Инструкция по использованию:'
                            f'<p>1) Экспортируем каталог с компьютерами из Active Directory в "workstations.txt".'
                            f'<p>2) Помещаем файл в директорию с программой.'
                            f'<p>3) Там должна находиться таблица "compare_sids.xlsx" с определенным шаблоном'
                            f'<p>4) Запускаем от имени администратора.'
                            f'<p>5) Запускаем "Получить SIDs" и ждем завершения.'
                            f'<p>6) Запускаем "Проверка на дублирование SIDs" и ждем завершения.'
                            f'<p>7) Вывод результатов производится в "compare_sids.xlsx"')
        self.close_button = QPushButton("Закрыть")

        layout.addWidget(self.label)
        layout.addWidget(self.close_button)

        self.close_button.clicked.connect(self.close)

        self.setLayout(layout)

    def center(self):
        parent_geometry = self.parent().geometry()
        self.setGeometry(
            parent_geometry.x() + parent_geometry.width() // 2 - self.width() // 2,
            parent_geometry.y() + parent_geometry.height() // 2 - self.height() // 2,
            self.width(),
            self.height()
        )


class CreateTableWindow(QDialog):
    def __init__(self, parent, title, x, y):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedWidth(x)
        self.setFixedHeight(y)
        self.setWindowIcon(QIcon('icon.ico'))
        self.center_window()

        layout = QVBoxLayout()
        result_path = 1

        result = funcs.compare_sids_table_exists()

        if result:
            self.label = QLabel(f'<center>Таблица "compare_sids.xlsx" уже существует. Перезаписать?</center>')
        else:
            funcs.create_compare_sids_table()
            result = funcs.compare_sids_table_exists()
            result_path = 2
            if result:
                self.label = QLabel(f'<center>Шаблон таблицы успешно создан!')
            else:
                self.label = QLabel(f'<center>Возникла проблема с созданием шаблона таблицы</center>')


        layout.addWidget(self.label)

        if result_path == 1:
            self.no_button = QPushButton("Нет")
            self.yes_button = QPushButton("Да")

            layout.addWidget(self.no_button)
            layout.addWidget(self.yes_button)

            self.yes_button.clicked.connect(self.close)
            self.no_button.clicked.connect(self.close)

        elif result_path == 2:
            self.ok_button = QPushButton("Ок")

            layout.addWidget(self.ok_button)

            self.ok_button.clicked.connect(self.close)

        layout.addLayout(layout)
        self.setLayout(layout)

    def center_window(self):
        parent_geometry = self.parent().geometry()
        self.setGeometry(
            parent_geometry.x() + parent_geometry.width() // 2 - self.width() // 2,
            parent_geometry.y() + parent_geometry.height() // 2 - self.height() // 2,
            self.width(),
            self.height()
        )


class CreateTableResult(QDialog):
    def __init__(self, parent, title, x, y):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setFixedWidth(x)
        self.setFixedHeight(y)
        self.setWindowIcon(QIcon('icon.ico'))
        self.center()

        layout = QVBoxLayout()

        self.close_button = QPushButton("Ок")

        result = funcs.compare_sids_table_exists()
        if result:
            self.label = QLabel(f'Шаблон таблицы успешно создан!')
        else:
            self.label = QLabel(f'Возникла проблема с созданием шаблона таблицы')

        layout.addWidget(self.label)
        layout.addWidget(self.close_button)

        self.close_button.clicked.connect(self.close)

        self.setLayout(layout)

    def center(self):
        parent_geometry = self.parent().geometry()
        self.setGeometry(
            parent_geometry.x() + parent_geometry.width() // 2 - self.width() // 2,
            parent_geometry.y() + parent_geometry.height() // 2 - self.height() // 2,
            self.width(),
            self.height()
        )


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dublicate SIDs Check")
        self.setGeometry(100, 100, 400, 500)
        self.setFixedWidth(400)
        self.setFixedHeight(500)
        self.setWindowIcon(QIcon('icon.ico'))

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()

        self.about_button = QPushButton("About")
        self.create_table_button = QPushButton('Создать "compare_sids.xlsx"')
        self.get_sids_button = QPushButton("Получить SIDs")
        self.get_failed_sids_button = QPushButton("Получить недостающие SIDs")
        self.get_dubles_button = QPushButton("Проверка на дублирование SIDs")
        self.message_output = QTextEdit()  # log field
        self.message_output.setReadOnly(True)

        layout.addWidget(self.about_button)
        layout.addWidget(self.create_table_button)
        layout.addWidget(self.get_sids_button)
        layout.addWidget(self.get_failed_sids_button)
        layout.addWidget(self.get_dubles_button)
        layout.addWidget(self.message_output)

        self.about_button.clicked.connect(self.about_window)
        self.create_table_button.clicked.connect(self.create_compare_sids_table_func)
        self.get_sids_button.clicked.connect(self.get_domain_n_local_sids)
        # self.get_failed_sids_button.clicked.connect(self.get_failed_domain_n_local_sids)
        self.get_dubles_button.clicked.connect(self.get_dubles)

        central_widget.setLayout(layout)


    def create_compare_sids_table_func(self):
        result = funcs.compare_sids_table_exists()
        print(result)
        if result:
            self.secondary_window = CreateTableWindow(self, 'Таблица уже существует', 450, 100)
            self.secondary_window.exec_()
        else:
            self.secondary_window = CreateTableWindow(self, 'Результат создания таблицы', 450, 100)
            self.secondary_window.exec_()


    def about_window(self):
        self.secondary_window = AboutWindow(self, 'About', 550, 450)
        self.secondary_window.exec_()


    # run different thread for getting sids
    def get_domain_n_local_sids(self):
        self.message_output.append(f'Идет получение Domain и Local SIDs.'
                                   f'<p>Результаты будут записаны в "compare_sids.xlsx" '
                                   f'после получения последнего SID.'
                                   f'<p>Убедитесь, что таблица "compare_sids.xlsx" не открыта.'
                                   f'<p>Пожалуйста, подождите.'
                                   f'<p>************************************************')
        self.get_sids_button.setEnabled(False)
        self.sids_thread = WorkerThread(self.run_get_domain_n_local_sids)
        self.sids_thread.partial_result_ready.connect(self.on_partial_sids_received)
        self.sids_thread.result_ready.connect(self.on_sids_received)
        self.sids_thread.start()

    # get sids for each ws
    def run_get_domain_n_local_sids(self):
        for ws_name_output, message, fail_output, local_sid_output, domain_sid_output in funcs.get_names_and_sids(
                ws_list, sheet, compare_sids_xlsx, ws_list_discr):
            yield (ws_name_output, message, fail_output, local_sid_output, domain_sid_output)

    # handling partial results and output logs
    def on_partial_sids_received(self, result):
        global n
        ws_name_output, message, fail_output, local_sid_output, domain_sid_output = result

        print(f'INFO: ws_list: {ws_list}')
        print(f'INFO: ws_list_discr: {ws_list_discr}')
        print(f'INFO: n: {n}')

        if message:
            self.message_output.append("<b>" + ws_name_output + ":</b>"
                                       + "<p><b>Name:</b> " + ws_list[n]
                                       + "<p><b>Discr:</b> " + ws_list_discr[n]
                                       + "<p><b>Local SID</b> - " + message
                                       + "<p><b>Domain SID</b> - " + domain_sid_output
                                       + "<p>************************************************")
        else:
            self.message_output.append("<b>" + ws_name_output + ":</b>"
                                       + "<p><b>Name:</b> " + ws_list[n]
                                       + "<p><b>Discr:</b> " + ws_list_discr[n]
                                       + "<p><b>Local SID</b> - " + local_sid_output
                                       + "<p><b>Domain SID</b> - " + domain_sid_output
                                       + "<p>************************************************")
        n += 1

    # handling full results
    def on_sids_received(self, results):
        global n
        self.message_output.append('Domain и Local SIDs получены и записаны в "compare_sids.xlsx"')
        self.message_output.append("************************************************")
        n = 0
        self.get_sids_button.setEnabled(True)

    def get_dubles(self):  # TODO fix it
        self.message_output.append("Идет поиск дублирующихся SIDs.")
        self.get_dubles_button.setEnabled(False)
        self.dubles_thread = WorkerThread(self.run_get_dubles)
        self.dubles_thread.partial_result_ready.connect(self.on_partial_dubles_received)
        self.dubles_thread.result_ready.connect(self.on_dubles_checked)
        self.dubles_thread.start()

    # run different thread for getting dubles
    def run_get_dubles(self):
        funcs.looking_for_dubles(sheet, compare_sids_xlsx, not_domain_sid=True)
        funcs.looking_for_dubles(sheet, compare_sids_xlsx, not_domain_sid=False)

    # handling partial results
    def on_partial_dubles_received(self):
        self.message_output.append('...')

    # handling full results
    def on_dubles_checked(self, results):
        self.message_output.append('Поиск дублей закончен. Результаты записаны в "compare_sids.xlsx"')
        self.message_output.append("************")
        self.get_dubles_button.setEnabled(True)

    # Метод для получения недостающих Domain и Local SIDs
    # def get_failed_domain_n_local_sids(self):
    #     self.message_output.append(f"Идет получение недостающих Domain и Local SIDs. Пожалуйста, подождите."
    #                                f"<p>************")
    #     self.get_failed_sids_button.setEnabled(False)
    #     self.failed_sids_thread = WorkerThread(self.run_get_failed_domain_n_local_sids)
    #     self.failed_sids_thread.partial_result_ready.connect(self.on_partial_failed_sids_received)
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
    #         self.message_output.append(ws_name_output + ":" + "<p>Name: " + ws_list[n] + "<p>Local SID - "
    #         + local_sid_output + '<p>Domain SID - ' + domain_sid_output + "<p>************")
    #
    # # Метод для обработки полного результата
    # def on_failed_sids_received(self, results):
    #     self.message_output.append('Domain и Local SIDs получены и записаны в "compare_sids.xlsx"')
    #     self.message_output.append("************")
    #     self.get_failed_sids_button.setEnabled(True)