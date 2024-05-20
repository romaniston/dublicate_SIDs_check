import openpyxl
import subprocess
import time
import funcs

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