import os.path
import time
import random
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
print('')
try:
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'credentials.json')
    print('Добрый день Александр')
    f = open('text.txt','r+')
    sum_in_file = f.read()
    print(sum_in_file)
    if len(sum_in_file) == 0:
        link = input('Вставьте ссылку на ваш проект:')
        f.write(link)
    elif len(sum_in_file) > 0:
        napisi = input('Подставить прошлую ссылку Y/N:').upper()
        if napisi == 'Y':
            link = sum_in_file
        else:
            link = input('Вставьте ссылку на ваш проект:')
            f = open('text.txt', 'w')
            f.write(link)
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    SAMPLE_SPREADSHEET_ID = f'{link[39:83]}'



    service = build('sheets', 'v4', credentials=credentials).spreadsheets().values()

    get1 =service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                       range='bot!B1', ).execute()
    get1 = get1.get('values')
    get1 = [element for row in get1 for element in row]
    column = get1[0]
    time.sleep(1)
    get1 =service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                       range='bot!B2', ).execute()
    get1 = get1.get('values')
    get1 = [element for row in get1 for element in row]
    columnf = get1[0]
    time.sleep(1)
    get1 =service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                       range='bot!B3', ).execute()
    get1 = get1.get('values')
    get1 = [element for row in get1 for element in row]
    columng = get1[0]
    time.sleep(1)
    get1 =service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                       range='bot!B4', ).execute()
    get1 = get1.get('values')
    get1 = [element for row in get1 for element in row]
    listname = get1[0]
    time.sleep(1)
    SAMPLE_RANGE_NAME = f'{listname}'
    get1 = service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                        range='bot!B5', ).execute()
    get1 = get1.get('values')
    get1 = [element for row in get1 for element in row]
    comparison = get1[0]

    result = service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                         range=SAMPLE_RANGE_NAME).execute()

    values = service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                         range=f'{listname}!{comparison}2:{comparison}{columng}').execute()

    data_from_sheet = result.get('values', [])
    values = values.get('values', [])[1:]

    values = [element for row in values for element in row]

    data = []
    number = []
    summa = []

    directory = f'{os.getcwd()}/FILES'
    brekfast = 0
    time_line = 0
    r = 0
    procent100 = 0.0
    procent1 = 0.0
    percent_1 = len(number) / 100
    jndeks = 0

    print(f'столбец для введения суммы:{column}')
    print(f'Столбец для введения комментария:{columnf}')
    print(f'Строка до которой стирать значения:{columng}')
    print(f'Название листа: {listname}')
    print(f'Столбец сравнения:{comparison}')
    a = input('Всё верно?(Y/N):')
    if a.upper() == 'Y':
        clear1 = service.clear(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                               range=f'{listname}!{column}2:{column}{columng}').execute()

        clear2 = service.clear(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                               range=f'{listname}!{columnf}2:{column}{columng}').execute()

        if int(columng) != len(values)+2:
            print(f'Предупреждение: В вашей таблице есть пустые столбцы в количестве->{int(columng) - len(values)-2}\nПрограмма будет работать некоректно просьба исправить')
            while True:
                pass
        else:
            for filename in os.listdir(directory):
                filenames = os.path.join(directory, filename)
                if os.path.isfile(filenames) and filename.endswith('.txt'):
                    file = open(f'{filenames}', 'r')
                    no_grups = ''

                    for i in file:
                        no_grups += i
                    grups = no_grups.split(';')
                    grups = list(filter(None, grups))
                    i = 0
                    while i < int(len(grups) - 12):
                        if i == 0:
                            data.append(grups[i])
                        else:
                            data.append(grups[i][1:])
                        i = i + 12
                    i = 5
                    while i < int(len(grups) - 12):
                        number.append(grups[i])
                        i = i + 12
                    i = 8
                    while i < int(len(grups)):
                        summa.append(grups[i])
                        i = i + 12
            procent1 = len(number) / 100
            for i in range(len(number)):
                try:
                    f = values.index(f'{number[i]}')
                    res2 = service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                       range=f'{listname}!{column}{f + 2}', ).execute()
                    time.sleep(1)
                    res5 = service.get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                       range=f'{listname}!{columnf}{f + 2}', ).execute()
                    time.sleep(1)
                    res4 = res2.get('values')
                    res6 = res5.get('values')
                    try:
                        res4 = [element for row in res4 for element in row]
                        res = service.update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                             range=f'{listname}!{column}{f + 2}',
                                             valueInputOption='RAW',
                                             body={'values': [
                                                 [float(summa[i].replace(",", ".")) + float(res4[0])]]}).execute()
                    except:
                        res = service.update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                             range=f'{listname}!{column}{f + 2}',
                                             valueInputOption='RAW',
                                             body={'values': [[float(summa[i].replace(",", "."))]]}).execute()
                    try:
                        res6 = [element for row in res6 for element in row]
                        res3 = service.update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                              range=f'{listname}!{columnf}{f + 2}',
                                              valueInputOption='RAW',
                                              body={'values': [[
                                                  f'{data[i].replace("-", "/")} {float(summa[i].replace(",", "."))} сб \n                                          {res6[0]}']]}).execute()
                        time.sleep(1)
                    except:
                        res3 = service.update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                              range=f'{listname}!{columnf}{f + 2}',
                                              valueInputOption='RAW',
                                              body={'values': [[
                                                  f'{data[i].replace("-", "/")} {float(summa[i].replace(",", "."))} сб']]}).execute()
                        time.sleep(1)
                        procent100 = i / procent1
                        print(f'{round(procent100, 2)}%')
                except:
                    print(f'Номера {number[i]} нет в таблице\nПерехожу на следущий номер,работа продолжается')
    else:
        print('Переделывай')
except Exception as err:
    print(err)
    while True:
        pass
