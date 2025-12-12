from PyQt5.QtWidgets import QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget, QTextEdit, QDialog, QHBoxLayout, QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QThread, pyqtSignal

from modules import funcs

from collections.abc import Iterable


ws_list, ws_list_discr = funcs.get_ws_names_and_discr()
n = 0

compare_sids_xlsx = None
sheet = None


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
        self.get_sids_button = QPushButton("Получить SIDs")
        self.get_failed_sids_button = QPushButton("Получить недостающие SIDs")
        self.get_dubles_button = QPushButton("Проверка на дублирование SIDs")
        self.message_output = QTextEdit()  # log field
        self.message_output.setReadOnly(True)

        layout.addWidget(self.about_button)
        layout.addWidget(self.get_sids_button)
        layout.addWidget(self.get_failed_sids_button)
        layout.addWidget(self.get_dubles_button)
        layout.addWidget(self.message_output)

        self.about_button.clicked.connect(self.about_window)
        self.get_sids_button.clicked.connect(self.check_and_get_sids)
        self.get_failed_sids_button.clicked.connect(self.get_failed_domain_n_local_sids)
        self.get_dubles_button.clicked.connect(self.get_dubles)

        central_widget.setLayout(layout)

        self.compare_sids_xlsx = None
        self.sheet = None


    def check_and_get_sids(self):
        if not funcs.compare_sids_table_exists():
            success = funcs.create_compare_sids_table()
            if success:
                self.message_output.append("Таблица 'compare_sids.xlsx' успешно создана!<p>")
                self.compare_sids_xlsx, self.sheet = funcs.table_var_assign()
                self.get_domain_n_local_sids()
            else:
                self.message_output.append("Ошибка: Не удалось создать таблицу!")
        else:
            reply = QMessageBox.question(
                self, 
                'Таблица уже существует',
                'Таблица "compare_sids.xlsx" уже существует. Создать новую таблицу (старая будет перезаписана)?',
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            
            if reply == QMessageBox.Yes:
                success = funcs.create_compare_sids_table()
                if success:
                    self.message_output.append("Таблица 'compare_sids.xlsx' успешно перезаписана!<p>")
                    self.compare_sids_xlsx, self.sheet = funcs.table_var_assign()
                    self.get_domain_n_local_sids()
                else:
                    self.message_output.append("Ошибка: Не удалось создать новую таблицу!")
            else:
                self.message_output.append("Операция отменена пользователем.")


    def about_window(self):
        self.secondary_window = AboutWindow(self, 'About', 550, 450)
        self.secondary_window.exec_()


    # run different thread for getting sids
    def get_domain_n_local_sids(self):
        # Проверяем что таблица загружена
        if self.compare_sids_xlsx is None or self.sheet is None:
            self.message_output.append("Ошибка: Таблица не загружена!")
            return
            
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
                ws_list, self.sheet, self.compare_sids_xlsx, ws_list_discr):
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

    def get_dubles(self):
            self.message_output.append("Идет поиск дублирующихся SIDs.")
            self.get_dubles_button.setEnabled(False)
            self.dubles_thread = WorkerThread(self.run_get_dubles)
            self.dubles_thread.partial_result_ready.connect(self.on_partial_dubles_received)
            self.dubles_thread.result_ready.connect(self.on_dubles_checked)
            self.dubles_thread.start()


    # run different thread for getting dubles
    def run_get_dubles(self):
        try:
            if not funcs.compare_sids_table_exists():
                success = funcs.create_compare_sids_table()
                if not success:
                    error_html = f"<font color='red'><b>Ошибка: Не удалось создать таблицу!</b></font>"
                    yield ("error", error_html)
                    return
            
            compare_sids_xlsx, sheet = funcs.table_var_assign()
            
            if compare_sids_xlsx is None or sheet is None:
                error_html = f"<font color='red'><b>Ошибка: Не удалось загрузить таблицу!</b></font>"
                yield ("error", error_html)
                return

            result = funcs.looking_for_dubles(sheet, compare_sids_xlsx, not_domain_sid=True)

            if isinstance(result, str) and "font color='red'" in result:
                yield ("error", result)
            elif result is None:
                yield ("local", "Поиск дублей Local SID завершён")
            elif isinstance(result, Iterable) and not isinstance(result, (str, bytes)):
                for item in result:
                    yield ("local", str(item))
            else:
                yield ("local", f"Результат Local SID: {result}")

            result = funcs.looking_for_dubles(sheet, compare_sids_xlsx, not_domain_sid=False)

            if isinstance(result, str) and "font color='red'" in result:
                yield ("error", result)
            elif result is None:
                yield ("domain", "Поиск дублей Domain SID завершён")
            elif isinstance(result, Iterable) and not isinstance(result, (str, bytes)):
                for item in result:
                    yield ("domain", str(item))
            else:
                yield ("domain", f"Результат Domain SID: {result}")

        except Exception as e:
            error_html = f"<font color='red'><b>Ошибка при поиске дубликатов: {str(e)}</b></font>"
            print(f"Ошибка в run_get_dubles: {e}")
            yield ("error", error_html)

    # handling partial results
    def on_partial_dubles_received(self, result):
        if isinstance(result, tuple) and len(result) == 2:
            if result[0] == "error":
                self.message_output.append(result[1])
            elif result[0] == "local" or result[0] == "domain":
                self.message_output.append(result[1])
        else:
            self.message_output.append('...')


    # handling full results
    def on_dubles_checked(self, results):
        self.message_output.append('Поиск дублей закончен. Результаты записаны в "compare_sids.xlsx"')
        self.message_output.append("************************************************")
        self.get_dubles_button.setEnabled(True)


    def get_failed_domain_n_local_sids(self):
        """Метод для получения недостающих SID"""
        if not funcs.compare_sids_table_exists():
            self.message_output.append("Таблица 'compare_sids.xlsx' не найдена. Сначала создайте таблицу или получите все SID.")
            return
        
        compare_sids_xlsx, sheet = funcs.table_var_assign()
        
        if compare_sids_xlsx is None or sheet is None:
            self.message_output.append("Ошибка: Не удалось загрузить таблицу!")
            return
    
        self.message_output.append(f"Идет получение недостающих Domain и Local SIDs. Пожалуйста, подождите."
                                f"<p>************")
        self.get_failed_sids_button.setEnabled(False)
        self.failed_sids_thread = WorkerThread(self.run_get_failed_domain_n_local_sids, compare_sids_xlsx, sheet)
        self.failed_sids_thread.partial_result_ready.connect(self.on_partial_failed_sids_received)
        self.failed_sids_thread.result_ready.connect(self.on_failed_sids_received)
        self.failed_sids_thread.start()


    def run_get_failed_domain_n_local_sids(self, compare_sids_xlsx, sheet):
        for ws_name_output, message, skip_output, local_sid_output, domain_sid_output in funcs.get_missing_sids(
                ws_list, sheet, compare_sids_xlsx, ws_list_discr):
            yield (ws_name_output, message, skip_output, local_sid_output, domain_sid_output)


    def on_partial_failed_sids_received(self, result):
        global n
        
        ws_name_output, message, skip_output, local_sid_output, domain_sid_output = result
        
        print(f'INFO: n = {n}, ws_list length = {len(ws_list)}, ws_list_discr length = {len(ws_list_discr)}')
        
        if skip_output == "skip":
            # if SID is here already, then pass it
            self.message_output.append(f"<b>{ws_name_output}:</b> SID уже получены, пропускаем."
                                    f"<p><b>Name:</b> {ws_list[n]}"
                                    f"<p><b>Discr:</b> {ws_list_discr[n]}"
                                    f"<p>************************************************")
        elif message:
            # if fail message is here
            self.message_output.append(f"<b>{ws_name_output}:</b>"
                                    f"<p><b>Name:</b> {ws_list[n]}"
                                    f"<p><b>Discr:</b> {ws_list_discr[n]}"
                                    f"<p><b>Local SID</b> - {message}"
                                    f"<p><b>Domain SID</b> - {domain_sid_output}"
                                    f"<p>************************************************")
        else:
            # if SID is here success
            self.message_output.append(f"<b>{ws_name_output}:</b>"
                                    f"<p><b>Name:</b> {ws_list[n]}"
                                    f"<p><b>Discr:</b> {ws_list_discr[n]}"
                                    f"<p><b>Local SID</b> - {local_sid_output}"
                                    f"<p><b>Domain SID</b> - {domain_sid_output}"
                                    f"<p>************************************************")
        
        n += 1  # Увеличиваем счетчик


    def on_failed_sids_received(self, results):
        global n
        
        self.message_output.append('Недостающие Domain и Local SIDs получены и записаны в "compare_sids.xlsx"')
        self.message_output.append("************************************************")
        n = 0
        self.get_failed_sids_button.setEnabled(True)