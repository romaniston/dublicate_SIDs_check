import openpyxl
import subprocess
import time
import asyncio

# Создаем переменные для работы с таблицей compare_sids_xlsx
def table_var_assign():
    try:
        compare_sids_xlsx = openpyxl.open('compare_sids.xlsx')
        sheet = compare_sids_xlsx.active
        compare_sids_xlsx.save('compare_sids.xlsx')
    except PermissionError:
        print('Нет доступа к таблице compare_sids.xlsx. Возможно, её нужно закрыть. Программа остановлена.')
        time.sleep(999999)
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
    result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='cp866')
    if 'число' or 'bytes' in str(result).split():
        bool_var = True
    else:
        bool_var = False
    return bool_var

# Выполняем cmd команду на получение SID (local или domain)
domain_sid_dict = {}
local_sid_dict = {}
async def complete_cmd_command(command, n, domain_sid=True):
    result = await asyncio.to_thread(subprocess.run, command, shell=True, capture_output=True, text=True, encoding='cp866')
    sid = str(result.stdout).split(':')
    clear_sid = sid[-1].strip()
    if domain_sid:
        domain_sid_dict[n] = clear_sid
    else:
        local_sid_dict[n] = clear_sid

# Получаем доменный SID
async def get_domain_sid(ws_list, n):
    await complete_cmd_command('cd c:\\ & cd pstools & psgetsid ' + str(ws_list[n]) + '$', n, True)

# Получаем локальный SID
async def get_local_sid(ws_list, n):
    await complete_cmd_command('cd c:\\ & cd pstools & psgetsid \\\\' + str(ws_list[n]), n, False)

# Получаем список из корутин для получения domain или local SID
def get_coroutine_list_sid(ws_list, domain_sid=True):
    coroutine_list = []
    if domain_sid:
        for n in range(len(ws_list)):
            coroutine_list.append(get_domain_sid(ws_list, n))
    else:
        for n in range(len(ws_list)):
            coroutine_list.append(get_local_sid(ws_list, n))
    return coroutine_list

# Запускаем event_loop для асинхронного получения списков sid
async def get_sid_in_list(ws_list, domain_sid=True):
    coroutine_list = get_coroutine_list_sid(ws_list, domain_sid)
    await asyncio.gather(*coroutine_list)
    if domain_sid:
        print(domain_sid_dict)
    else:
        print(local_sid_dict)

# Запись данных WS и получение Local и Domain SID
def get_names_and_sids(ws_list, sheet, compare_sids_xlsx, ws_list_discr):

    event_loop = asyncio.get_event_loop()

    n = 0
    for string in range(len(ws_list)):

        done_discr_fulled = ''
        done_discr = ''

        print(f'Workstation #{1+n}')
        ping_or_not_bool = ping_or_not(ws_list, n)
        sheet[2 + (string)][0].value = 1 + n
        sheet[2 + (string)][1].value = ws_list[0 + n]

        # Заполняем столбец описания WS
        if ws_list_discr[0 + n] != []:
            for done_discr in ws_list_discr[0 + n]:
                done_discr_fulled += ' ' + done_discr

            sheet[2 + (string)][2].value = done_discr_fulled.strip()
            print(f'Описание WS: {done_discr_fulled.strip()}')
        else:
            print('Описания нет')

        if not ping_or_not_bool:
            sheet[2 + (string)][3].value = 'Workstation не пингуется'
            print('Workstation не пингуется. Local SID не получен.')
        else:
            if get_local_sid(ws_list, 0 + n)[0] != 'S':
                sheet[2 + (string)][3].value = 'Не удалось получить Local SID. Возможно утилита запущена не от имени администратора.'
                print('Локальный SID: Не удалось получить Local SID. Возможно утилита запущена не от имени администратора')
            else:
                event_loop.run_until_complete()

                local_sid = get_local_sid(ws_list, 0 + n)
                sheet[2 + (string)][3].value = local_sid
                print(f'Локальный SID: {local_sid}')
        domain_sid = get_domain_sid(ws_list, 0 + n)
        sheet[2 + (string)][4].value = domain_sid
        print(f'Доменный SID: {domain_sid}')
        n += 1
        print('--------------')
        compare_sids_xlsx.save('compare_sids.xlsx')

# Функция для поиск дублей Local и Domain SID
def check_dubles(ws_list, sheet, compare_sids_xlsx, not_domain_sid):

    # При True ищется Local SID, при False - Domain SID
    if not_domain_sid:
        print('Идёт поиск дублей Local SID...')
        domain_sid = 0
    else:
        print('Идёт поиск дублей Domain SID...')
        domain_sid = 1

    is_header_field = True

    # Поиск дублей
    for string in range(len(ws_list) + 1):
        if is_header_field:
            is_header_field = False
            continue

        compared_ws_num_source = sheet[1 + string][0].value
        sid_val_source = sheet[1 + string][3 + domain_sid].value
        list_of_compared = []

        for compare_string in range(len(ws_list) + 1):
            compared_ws_num = sheet[1 + compare_string][0].value
            compared_ws_name = sheet[1 + compare_string][1].value
            sid_val = sheet[1 + compare_string][3 + domain_sid].value
            if compared_ws_num_source != compared_ws_num:
                if sid_val == sid_val_source:
                    if sid_val != 'Workstation не пингуется':
                        if sid_val != 'Не удалось получить Local SID. Возможно утилита запущена не от имени администратора.':
                            list_of_compared.append(f'№{compared_ws_num} - {compared_ws_name}')
                        else:
                            pass

        list_of_compared_done = ''
        if list_of_compared != []:
            for duble in list_of_compared:
                list_of_compared_done += str(duble) + ", "
            sheet[1 + string][5 + domain_sid].value = f'Найдены дубли: {list_of_compared_done}'
        else:
            sheet[1 + string][5 + domain_sid].value = 'Совпадений нет'

    try:
        compare_sids_xlsx.save('compare_sids.xlsx')
    except PermissionError:
        print('Нет доступа к таблице compare_sids.xlsx. Возможно, её нужно закрыть. Программа остановлена.')
        time.sleep(999999)