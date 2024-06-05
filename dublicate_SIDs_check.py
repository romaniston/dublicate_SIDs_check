import openpyxl
import subprocess
import time
import funcs
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel,\
    QVBoxLayout, QWidget, QTextEdit, QDialog, QVBoxLayout
from PyQt5.QtGui import QIcon

# Создаем переменные для работы с таблицей compare_sids_xlsx
compare_sids_xlsx, sheet = funcs.table_var_assign()

# Получаем массивы с именами и описаниями WS из workstations.txt
ws_list, ws_list_discr = funcs.get_ws_names_and_discr()

# Создаем класс диалогового окна
class SecondaryWindow(QDialog):
    def __init__(self, title, x, y):
        super().__init__()
        self.setWindowTitle(title)
        self.setGeometry(150, 150, x, y)
        self.setFixedWidth(x)
        self.setFixedHeight(y)
        self.setWindowIcon(QIcon('icon.ico'))

        # Создаем вертикальный layout
        layout = QVBoxLayout()

        # Добавляем текстовое поле
        self.label = QLabel(f"<p><h3>Dublicate SIDs check предназначен для проверки локальной сети"
                            f"<p>на предмет дублирования Domain и Local SIDs.</h3>"
                            f"<p>************"
                            f"<p>Инструкция по использованию:"
                            f"<p>1) Выгружаем из Active Directory .txt файл с описанием рабочих станций не внося в него правок."
                            f"<p>2) Файл должен называться 'workstations.txt'. Помещаем его в директорию со скриптом."
                            f"<p>3) Также в той же директории должна находится таблица 'compare_sids.xlsx' с определенным шаблоном"
                            f"<p>4) Запускаем скрипт от имени администратора. Нажимаем кнопку 'Сканировать все WS и получить Domain и Local SIDs'"
                            f"<p>5) По окончанию операции в таблицу 'compare_sids.xlsx' будут записаны все Domain и Local SIDs из списка которые удалось получить"
                            f"<p>6) Нажимаем кнопку 'Запуск поиска дублей Domain и Local SIDs'"
                            f"<p>7) По окончании операции в таблицу будет выведена информация о дублирующихся SIDs в специальные столбцы")
        layout.addWidget(self.label)

        # Добавляем кнопку для закрытия окна
        self.close_button = QPushButton("Закрыть")
        layout.addWidget(self.close_button)

        # Подключаем сигнал к слоту для закрытия окна
        self.close_button.clicked.connect(self.close)

        # Устанавливаем layout
        self.setLayout(layout)

# Создаем класс основного окна
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dublicate SIDs Check")
        self.setGeometry(100, 100, 400, 500)  # Устанавливаем размеры окна
        self.setFixedWidth(400)
        self.setFixedHeight(500)
        self.setWindowIcon(QIcon('icon.ico'))

        # Создаем центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Создаем вертикальный layout
        layout = QVBoxLayout()

        # Добавляем кнопку 1
        self.help_for_use_button = QPushButton("Как пользоваться")
        layout.addWidget(self.help_for_use_button)

        # Добавляем кнопку 2
        self.get_sids_button = QPushButton("Сканировать все WS и получить Domain и Local SIDs")
        layout.addWidget(self.get_sids_button)

        # Добавляем кнопку 3
        self.get_dubles_button = QPushButton("Запуск поиска дублей Domain и Local SIDs")
        layout.addWidget(self.get_dubles_button)

        # Добавляем текстовое поле для вывода сообщений
        self.message_output = QTextEdit()
        self.message_output.setReadOnly(True)  # Делаем поле только для чтения
        layout.addWidget(self.message_output)

        # Устанавливаем layout для центрального виджета
        central_widget.setLayout(layout)

        # Присваиваем событие нажатия на кнопку к определенной функции и определяем вызов определенного слота
        self.get_sids_button.clicked.connect(self.get_domain_n_local_sids)
        self.help_for_use_button.clicked.connect(self.help_for_use)
        self.get_dubles_button.clicked.connect(self.get_dubles)

    def help_for_use(self):
        self.secondary_window = SecondaryWindow('Как пользоваться', 720, 350)
        self.secondary_window.exec_()

    def get_domain_n_local_sids(self):
        self.message_output.append(f"Идет получение Domain и Local SIDs. Пожалуйста, подождите."
                                   f"<p>************")
        self.get_sids_button.setEnabled(False)
        QApplication.processEvents()
        for ws_name_output, message, fail_output, local_sid_output, domain_sid_output in funcs.get_names_and_sids(
                ws_list, sheet, compare_sids_xlsx, ws_list_discr):
            if message:
                self.message_output.append(ws_name_output + ":" + "<p>Local SID - " + message +
                                           '<p>Domain SID - ' + domain_sid_output + "<p>************")
            else:
                self.message_output.append(ws_name_output + ":" + "<p>Local SID - " + local_sid_output +
                                           '<p>Domain SID - ' + domain_sid_output + "<p>************")
            QApplication.processEvents()
        self.message_output.append('Domain и Local SIDs получены и записаны в "compare_sids.xlsx"')
        self.message_output.append("************")
        self.get_sids_button.setEnabled(True)

    def get_dubles(self):
        self.message_output.append("Идет поиск дублирующихся SIDs.")
        self.get_dubles_button.setEnabled(False)  # Делаем кнопку неактивной
        funcs.check_dubles(ws_list, sheet, compare_sids_xlsx, True)
        funcs.check_dubles(ws_list, sheet, compare_sids_xlsx, False)
        self.message_output.append('Поиск закончен. Результаты записаны в "compare_sids.xlsx"')
        self.message_output.append("************")
        self.get_dubles_button.setEnabled(True)  # Делаем кнопку активной

# Эта часть кода позволяет запускать приложение
app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec_())









# Создаем переменные для работы с таблицей compare_sids_xlsx
compare_sids_xlsx, sheet = funcs.table_var_assign()

infinite_run = True

while infinite_run == True:

    run = input(f'Выберите действие:'
                f'\n1) Сканирование всех Workstations в сети и получение'
                f'\n    Local SID и Domain SID с записью в файл'
                f'\n2) Сканирование таблицы на предмет дублирующихся'
                f'\n    Local SID и Domain SID с записью в файл'
                f'\n>>> ')

    # Получаем массивы с именами и описаниями WS из workstations.txt
    ws_list, ws_list_discr = funcs.get_ws_names_and_discr()
    if run == '1':

        funcs.get_names_and_sids(ws_list, sheet, compare_sids_xlsx, ws_list_discr)

        print(f'--------------'
              f'\nОперация закончена'
              f'\n--------------')

    elif run == '2':

        funcs.check_dubles(ws_list, sheet, compare_sids_xlsx, True)
        funcs.check_dubles(ws_list, sheet, compare_sids_xlsx, False)

        print(f'--------------'
              f'\nОперация закончена. Файл "compare_sids.xlsx" обновлен.'
              f'\n--------------')