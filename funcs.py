import openpyxl
import subprocess
import time
from openpyxl.styles import PatternFill

output_message_ws_no_ping = 'Workstation не пингуется. Local SID не получен.'
output_message_dont_get_local_sid = 'Не удалось получить Local SID. Возможно утилита запущена не от имени администратора.'

# Создаем переменные для работы с таблицей compare_sids_xlsx
def table_var_assign():
    compare_sids_xlsx = openpyxl.open('compare_sids.xlsx')
    sheet = compare_sids_xlsx.active
    compare_sids_xlsx.save('compare_sids.xlsx')
    return compare_sids_xlsx, sheet

# Чтение из workstations.txt имен ПК и их описание и запись их в массивы
def get_ws_names_and_discr():
    ws_list = []
    ws_list_discr = []
    with open("workstations.txt", "r") as file:
        not_first_str = False
        for string in file:
            if not_first_str:
                split_string = string.split()
                ws_list.append(split_string[0])
                ws_list_discr.append(split_string[2:])
            else:
                pass
                not_first_str = True
    return ws_list, ws_list_discr

# Функция для проверки пингуется ли WS
def ping_or_not(ws_list, n):
    command = f'ping {ws_list[0 + n]}'
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='cp866', timeout=30)
        if 'число' in str(result).split():
            bool_var = True
        else:
            bool_var = False
    except subprocess.TimeoutExpired:
        print(f'Команда "{command}" превысила тайм-аут 30 секунд.')
        bool_var = False
    return bool_var

# Получаем доменный SID
def get_domain_sid(ws_list, ws_num):
    command = 'cd c:\\ & cd pstools & psgetsid ' + str(ws_list[ws_num]) + '$'
    result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='cp866')
    get_sid = str(result).split(':')
    get_clear_domain_sid = str(get_sid[-1]).split('\\n')
    return get_clear_domain_sid[1]

# Получаем локальный SID
def get_local_sid(ws_list, ws_num):
    command = 'cd c:\\ & cd pstools & psgetsid \\\\' + str(ws_list[ws_num])
    result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='cp866')
    time.sleep(5)
    get_sid = str(result).split(':')
    get_clear_local_sid = str(get_sid[-1]).split('\\n')
    return get_clear_local_sid[1]

# Запись данных WS и получение Local и Domain SID
def get_names_and_sids(ws_list, sheet, compare_sids_xlsx, ws_list_discr):
    n = 0
    for string in range(len(ws_list)):
        done_discr_fulled = ''
        done_discr = ''
        ws_name_output = f'Workstation #{1+n}'
        print('test')
        ping_or_not_bool = ping_or_not(ws_list, n)
        print('success')
        sheet[2 + (string)][0].value = 1 + n
        sheet[2 + (string)][1].value = ws_list[0 + n]
        print(1)
        for done_discr in ws_list_discr[0 + n]:
            done_discr_fulled += ' ' + done_discr
        print(2)
        sheet[2 + (string)][2].value = done_discr_fulled
        print(3)
        if not ping_or_not_bool:
            print(3.1)
            sheet[2 + (string)][3].value = output_message_ws_no_ping
            print(3.2)
            domain_sid = get_domain_sid(ws_list, 0 + n)
            yield ws_name_output, output_message_ws_no_ping, None, local_sid, domain_sid
            print(3.3)
        else:
            local_sid = get_local_sid(ws_list, 0 + n)
            domain_sid = get_domain_sid(ws_list, 0 + n)
            if local_sid[0] != 'S':
                sheet[2 + (string)][3].value = output_message_dont_get_local_sid
                yield ws_name_output, output_message_dont_get_local_sid, None, local_sid, domain_sid
            else:
                sheet[2 + (string)][3].value = local_sid
                yield ws_name_output, None, None, local_sid, domain_sid
            compare_sids_xlsx.save('compare_sids.xlsx')
        print(4)
        sheet[2 + (string)][4].value = domain_sid
        n += 1
        compare_sids_xlsx.save('compare_sids.xlsx')

# Цвет ячейки excel
def field_color(color):
    red_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

# Функция для поиск дублей Local и Domain SID
def check_dubles(sheet, compare_sids_xlsx, not_domain_sid):
    # Подсчет WS в compare_sids.xlsx
    ws_count = 0
    n = 1
    while sheet[n][0].value != None:
        ws_count += 1
        n += 1

    print(ws_count)

    # При True ищется Local SID, при False - Domain SID
    if not_domain_sid:
        print('Идёт поиск дублей Local SID...')
        domain_sid = 0
    else:
        print('Идёт поиск дублей Domain SID...')
        domain_sid = 1

    is_header_field = True

    # Поиск дублей
    for string in range(ws_count):
        if is_header_field:
            is_header_field = False
            continue

        compared_ws_num_source = sheet[1 + string][0].value
        sid_val_source = sheet[1 + string][3 + domain_sid].value
        list_of_compared = []

        for compare_string in range(ws_count):
            compared_ws_num = sheet[1 + compare_string][0].value
            compared_ws_name = sheet[1 + compare_string][1].value
            sid_val = sheet[1 + compare_string][3 + domain_sid].value
            if compared_ws_num_source != compared_ws_num:
                if sid_val == sid_val_source:
                    if sid_val != output_message_ws_no_ping:
                        if sid_val != output_message_dont_get_local_sid:
                            list_of_compared.append(f'№{compared_ws_num} - {compared_ws_name}')
                        else:
                            pass

        local_sid_field = sheet[1 + string][3].value
        list_of_compared_done = ''
        if list_of_compared != []:
            for duble in list_of_compared:
                list_of_compared_done += str(duble) + ", "
            sheet[1 + string][5 + domain_sid].value = f'Найдены дубли: {list_of_compared_done}'
            sheet[1 + string][5 + domain_sid].fill = field_color("FF0000")
        else:
            sheet[1 + string][5 + domain_sid].value = 'Совпадений нет'
            sheet[1 + string][5 + domain_sid].fill = field_color("00FF00")

        if local_sid_field[0] != 'S':
            sheet[1 + string][5].value = 'Local SID не получен'
            sheet[1 + string][5].fill = field_color("FF8000")
        else:
            pass

    try:
        compare_sids_xlsx.save('compare_sids.xlsx')
    except PermissionError:
        print('Нет доступа к таблице compare_sids.xlsx. Возможно, её нужно закрыть. Программа остановлена.')
        time.sleep(999999)