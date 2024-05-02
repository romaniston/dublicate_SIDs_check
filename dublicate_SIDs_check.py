import openpyxl
import subprocess
import time

infinite_run = True

while infinite_run == True:

    run = input(f'Выберите действие:'
                f'\n1) Сканирование всех Workstations в сети и получение'
                f'\n    Local SID и Domain SID с записью в файл'
                f'\n2) Сканирование таблицы на предмет дублирующихся'
                f'\n    Local SID и Domain SID с записью в файл'
                f'\n>>> ')

    # Чтение из workstations.txt имен ПК и запись их в массив
    def get_ws_names():
        ws_list = []
        with open("workstations.txt", "r") as file:
            not_first_str = False
            for string in file:
                if not_first_str:
                    ws_name_only = string.split()
                    ws_list.append(ws_name_only[0])
                else:
                    pass
                    not_first_str = True
        return ws_list

    # Получаем массив с именами WS из workstations.txt
    ws_list = get_ws_names()

    # Создаем переменные для работы с таблицей compare_sids_xlsx
    try:
        compare_sids_xlsx = openpyxl.open('compare_sids.xlsx')
        sheet = compare_sids_xlsx.active
    except PermissionError:
        print('Нет доступа к таблице compare_sids.xlsx. Возможно, её нужно закрыть. Программа остановлена.')
        time.sleep(999999)


    if run == 1:

        # Функция для проверки пингуется ли WS
        def ping_or_not(n):
            command = f'ping {ws_list[0 + n]}'
            result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='cp866')
            if 'число' in str(result).split():
                bool_var = True
            else:
                bool_var = False
            return bool_var

        # Получаем доменный SID
        def get_domain_sid(ws_num):
            command = 'cd c:\\ & cd pstools & psgetsid ' + str(ws_list[ws_num]) + '$'
            result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='cp866')
            get_sid = str(result).split(':')
            get_clear_domain_sid = str(get_sid[-1]).split('\\n')
            return get_clear_domain_sid[1]

        # Получаем локальный SID
        def get_local_sid(ws_num):
            # command = 'cd c:\\ & cd pstools & psexec \\\\' + str(ws_list[ws_num]) + ' -u kmpp\\администратор -p *** -accepteula -c psgetsid.exe'
            command = 'cd c:\\ & cd pstools & psgetsid \\\\' + str(ws_list[ws_num])
            result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='cp866')
            time.sleep(5)
            get_sid = str(result).split(':')
            get_clear_local_sid = str(get_sid[-1]).split('\\n')
            return get_clear_local_sid[1]

        # Запись данных в compare_sids.xlsx
        n = 0
        for string in range(len(ws_list)):
            try:
                print(f'Workstation #{1+n}')
                ping_or_not_bool = ping_or_not(n)
                sheet[2 + (string)][0].value = 1 + n
                sheet[2 + (string)][1].value = ws_list[0 + n]
                if not ping_or_not_bool:
                    sheet[2 + (string)][2].value = 'Workstation не пингуется'
                    print('Workstation не пингуется. Local SID не получен.')
                # if get_local_sid(0 + n)[0] != 'S':
                #     sheet[2 + (string)][2].value = 'Не удалось получить Local SID'
                #     print('Не удалось получить Local SID')
                else:
                    if get_local_sid(0 + n)[0] != 'S':
                        sheet[2 + (string)][2].value = 'Не удалось получить Local SID. Возможно утилита запущена не от имени администратора.'
                        print('Локальный SID: Не удалось получить Local SID. Возможно утилита запущена не от имени администратора.')
                    else:
                        local_sid = get_local_sid(0 + n)
                        sheet[2 + (string)][2].value = local_sid
                        print(f'Локальный SID: {local_sid}')
                domain_sid = get_domain_sid(0 + n)
                sheet[2 + (string)][3].value = domain_sid
                print(f'Доменный SID: {domain_sid}')
                sheet[2 + (string)][4].value = 'result'
                n += 1
                print('--------------')
                compare_sids_xlsx.save('compare_sids.xlsx')
            except Exception:
                print(f'Возникла неизвестная ошибка на итерации {0 + n}. Итерация пропущена.')
                pass

        print(f'--------------'
              f'\nОперация закончена'
              f'\n--------------')

    elif run == '2':

        # Функция для поиск дублей Local SID
        def get_dubles(not_domain_sid=True):

            # При True ищется Local SID, при False - Domain SID
            if not_domain_sid:
                print('Идёт поиск дублей Local SID...')
                domain_sid = 0
            else:
                print('Идёт поиск дублей Domain SID...')
                domain_sid = 1

            # Поиск дублей
            for string in range(len(ws_list) + 1):
                compared_ws_num_source = sheet[2 + string - 1][0].value
                compared_ws_name_source = sheet[2 + string - 1][1].value
                sid_val_source = sheet[2 + string - 1][2 + domain_sid].value
                dict_of_compared = {}

                for compare_string in range(len(ws_list) + 1):
                    compared_ws_num = sheet[2 + compare_string - 1][0].value
                    compared_ws_name = sheet[2 + compare_string - 1][1].value
                    sid_val = sheet[2 + compare_string - 1][2 + domain_sid].value
                    if compared_ws_num_source != compared_ws_num:
                        if sid_val == sid_val_source:
                            if sid_val != 'Workstation не пингуется':
                                if sid_val != 'Не удалось получить Local SID. Возможно утилита запущена не от имени администратора.':
                                    dict_of_compared['№' + str(compared_ws_num) + ' ' + str(compared_ws_name)] = sid_val
                                else:
                                    pass

                if dict_of_compared != {}:
                    sheet[2 + string - 1][4 + domain_sid].value = f'Найдены дубли: {dict_of_compared}'
                else:
                    pass

            try:
                compare_sids_xlsx.save('compare_sids.xlsx')
            except PermissionError:
                print('Нет доступа к таблице compare_sids.xlsx. Возможно, её нужно закрыть. Программа остановлена.')
                time.sleep(999999)

        get_dubles(True)
        get_dubles(False)

        print(f'--------------'
              f'\nОперация закончена'
              f'\n--------------')