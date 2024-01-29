import pandas as pd
import re
from datetime import datetime, timedelta
import json
import requests
import pytz
import time
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side
from settings import TOKEN, TG_CH, PARAMETERS

'''Рабочие значения токена'''
TOKEN = TOKEN
TG_CH = TG_CH
URL = 'https://api.telegram.org/bot'

def send_message(text,mod=0):
    global last_message_id
    ''' {0:'Print_new_massage', 1:'Edit_massage'} '''

    if mod == 1:
        # Редактируем последнее отправленное сообщение
        res = requests.post('https://api.telegram.org/bot{}/editMessageText'.format(TOKEN),
                            data={'chat_id': TG_CH, 'message_id': last_message_id, 'text': text})
    if mod == 0:
        # last_message_id = None 
        # Отправляем новое сообщение
        res = requests.get('https://api.telegram.org/bot{}/sendMessage'.format(TOKEN),
                           params={'chat_id': TG_CH, 'text': text})
        last_message_id = res.json().get('result', {}).get('message_id')


PARAMETERS = PARAMETERS

with open('right_instumets.json', 'r', encoding='utf-8') as f:
    RIGHT_INSTRUMENTS = json.load(f)

'ex = excel = dataframe'
def processing_excel_data(dataframe_name :str )->pd.DataFrame:
    '''Данная функция обрабатывает Excel файл. Удаляет пропуски, проверяет на то, что все иеди по аремени завершилис и приводит к значения к численным данным'''
    ex = pd.read_excel(dataframe_name)
    ex.dropna(inplace=True)
    pattern = r'\s*\(UTC\+\d+\)'
    ex['CREATE TIME'] = ex['CREATE TIME'].apply(lambda row: (re.sub(pattern, '', row)))
    ex['EXPIRATION TIME'] = ex['EXPIRATION TIME'].apply(lambda row:(re.sub(pattern, '', row)))
    # Преобразуем столбец 'EXPIRATION TIME' в объекты datetime
    ex['Help_time'] = pd.to_datetime(ex['EXPIRATION TIME'], format="%d.%m.%Y %H:%M:%S")
    # Добавляем столбец 'mask'
    ex['mask'] = datetime.now() >= ex['Help_time']
    # Фильтруем DataFrame по условию
    ex = ex.loc[ex['mask']]
    # Удаляем столбец 'mask', если он больше не нужен
    ex = ex.drop(columns=['mask','Help_time','STATUS', 'DESCRIPTION TARGET','DESCRIPTION BACKGROUND'])
    ex['INSTRUMENT'] = ex['INSTRUMENT'].apply(lambda x: x.split('.')[0])
    ex[["STOP LOSS", "PRICE", "TAKE PROFIT"]] = ex[["STOP LOSS", "PRICE", "TAKE PROFIT"]].apply(pd.to_numeric, errors='coerce')
    ex = ex.reset_index(drop=True)
    return ex
    '''Все корректно работает'''

def load_data(symbol_id:str, CREATE_TIME:str, EXPIRATION_TIME:str)->list:
    def transfrm_to_TIMESTAMP(date_string:str)->str:
        date = datetime.strptime(date_string, '%d.%m.%Y %H:%M:%S')
        return int(date.timestamp()) * 1000 
    start_time =  transfrm_to_TIMESTAMP(CREATE_TIME)
    end_time = transfrm_to_TIMESTAMP(EXPIRATION_TIME)
    duration = 60 
    size = int(round(((end_time - start_time)/1000/duration), 1))
    url = f'https://api-live.exante.eu/md/2.0/ohlc/{symbol_id}/{duration}'
    params = {
    'from': start_time, 
    'to': end_time, 
    'size': size,
    'type': 'quotes'}
    response = requests.get(url, auth=(PARAMETERS["API Key"], PARAMETERS["Secret Key"]), params=params)
    data = json.loads(response.text)
    return data


# class Process_manager:
#     def __init__(self, dataframe_name:str):
#         self.df = processing_excel_data(dataframe_name)
        
def main_data_processing(FILE_NAME:str):
    df = processing_excel_data(FILE_NAME)
    strategy_result = []
    send_message(f'Обработанно: 0% \nОставшееся время обработки - {len(df)} мин',mod=0)
    for index, row in df.iterrows():
        INSTRUMENT, SIDE, OPEN_ORDER_TYPE, STOP_LOSS, PRICE, TAKE_PROFIT, CREATE_TIME, EXPIRATION_TIME = row[['INSTRUMENT', 'SIDE', 
                                                                                                            'OPEN ORDER TYPE', 'STOP LOSS', 
                                                                                                            'PRICE','TAKE PROFIT', 
                                                                                                            'CREATE TIME',
                                                                                                            'EXPIRATION TIME']]
        if INSTRUMENT in RIGHT_INSTRUMENTS:
            symbol_id = RIGHT_INSTRUMENTS[INSTRUMENT]
            data = load_data(symbol_id, CREATE_TIME, EXPIRATION_TIME)
            result = Strategy_Tester(data, SIDE, OPEN_ORDER_TYPE, STOP_LOSS, PRICE, TAKE_PROFIT)
            strategy_result.append(result)
            time.sleep(60)
            
        else:
            strategy_result.append(('Данного символа нет в списке инструментов (right_instumets.json)',))
        send_message(f'Обработанно: {(index+1)/len(df)*100:.1f}% \nОставшееся время обработки - {len(df) - index - 1} мин',mod=1)
    df3 = pd.DataFrame(strategy_result)
    new_columns = ['Result', 'CLOSE RATE', 'PL', 'OPEN TIME', 'CLOSE TIME', 'DURATION (H)', 'OPEN TIME IN MLSK', 'CLOSE TIME IN MLSK' ]
    # Переименование столбцов
    df3.columns = new_columns
    f = pd.concat([df, df3], axis=1)
    f.to_excel(FILE_NAME, index=False)
                
            
            
def Strategy_Tester(data:list, side:str, open_order_type:str, stop_loss:float, price:float, take_profit:float)->tuple:
    df = pd.DataFrame(data)
    # Преобразование timestamp в формат datetime с учетом часового пояса UTC
    df['timestamp2'] = pd.to_datetime(df['timestamp'], unit='ms').dt.tz_localize('UTC')
    # Преобразуем время в часовой пояс UTC+4
    df['timestamp_utc4'] = df['timestamp2'].dt.tz_convert(pytz.timezone('Etc/GMT-3'))
    # Преобразуем дату в формат "d-m-Y H:M:S" и создадим новый столбец 'formatted_date'
    df['formatted_date'] = df['timestamp_utc4'].dt.strftime('%d-%m-%Y %H:%M:%S')
    df = df.drop(columns=["timestamp2",	"timestamp_utc4"])
    df[["close", "high", "low", "open"]] = df[["close", "high", "low", "open"]].apply(pd.to_numeric, errors='coerce')
    df = df[::-1]
    df = df.reset_index(drop=True)
    '''Алгоритм проверки'''
    new_df = df.copy()
    PRICE_INDICATOR = False
    STOP_LOSS_INDICATOR = False
    TAKE_PROFIT_INDICATOR = False
    # 1. Проверяем вход по цене
    if side == 'Buy':
        Flag=True
        if open_order_type == 'Limit':
            price_open = new_df[new_df['low'] <= price]
        else:
            price_open = new_df[new_df['high'] >= price]
    else:
        Flag = False
        if open_order_type == 'Limit':
            price_open = new_df[new_df['high'] >= price]
        else:
            price_open = new_df[new_df['low'] <= price]

    if not price_open.empty:
        index_price = price_open.index[0]
        new_df = new_df.iloc[index_price:].reset_index(drop=True)
        PRICE_INDICATOR = True
        approximate_time_enter, accurate_time_enter = price_open['formatted_date'].iloc[0], price_open['timestamp'].iloc[0]
    # 2. Проверяем стоп-лосс
    if Flag:
        stop_loss_res = new_df[new_df['low'] <= stop_loss]
    else:
        stop_loss_res = new_df[new_df['high'] >= stop_loss]

    if not stop_loss_res.empty:
        index_stop_loss = stop_loss_res.index[0]
        new_df = new_df.iloc[:index_stop_loss + 1].reset_index(drop=True)
        STOP_LOSS_INDICATOR = True
        approximate_time_ST, accurate_time_ST = stop_loss_res['formatted_date'].iloc[0], stop_loss_res['timestamp'].iloc[0]

    # 3. Проверяем на take profit
    if Flag:
        take_profit_res = new_df[new_df['high'] >= take_profit]
    else:
        take_profit_res = new_df[new_df['low'] <= take_profit]
        
    if not take_profit_res.empty:
        take_profit_index = take_profit_res.index[0]
        TAKE_PROFIT_INDICATOR = True
        approximate_time_TP, accurate_time_TP = take_profit_res['formatted_date'].iloc[0], take_profit_res['timestamp'].iloc[0]

    def timedelta(str1, str2):
        date_format = '%d-%m-%Y %H:%M:%S'
        date1 = datetime.strptime(str1, date_format)
        date2 = datetime.strptime(str2, date_format)
        # Вычисление разницы между датами
        time_difference = date2 - date1 
        # Извлечение часов, минут и секунд из разницы
        hours, remainder = divmod(time_difference.seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        # Форматирование результатов
        result_string = '{:02}:{:02}:{:02}'.format(hours, minutes, seconds)
        return result_string
        
    # 4. Вывод
    if PRICE_INDICATOR:
        if TAKE_PROFIT_INDICATOR:
            return("TP",take_profit, abs(take_profit-price), approximate_time_enter, approximate_time_TP, timedelta(approximate_time_enter,approximate_time_TP), accurate_time_enter, accurate_time_TP)
        elif STOP_LOSS_INDICATOR and not TAKE_PROFIT_INDICATOR:
            return ("SL",stop_loss, -1*abs(stop_loss-price), approximate_time_enter, approximate_time_ST, timedelta(approximate_time_enter,approximate_time_ST), accurate_time_enter, accurate_time_ST)
        else:
            last_price = new_df['close'].iloc[-1]
            approximate_time_ST = new_df['formatted_date'].iloc[-1]
            accurate_time_ST = new_df['timestamp'].iloc[-1]
            return ("EXP", 
                    last_price, 
                    last_price - price if side == 'Buy'else price - last_price, 
                    approximate_time_enter, 
                    approximate_time_ST, 
                    timedelta(approximate_time_enter,approximate_time_ST), 
                    accurate_time_enter, accurate_time_ST)
    else:
        return('NOT',)
        


def vizual(FILE_NAME:str):
    # Загружаете книгу Excel
    workbook = load_workbook(FILE_NAME)
    # Выбираете нужный лист (в данном случае, первый лист)
    sheet = workbook.active
    # Проходитесь по строкам, начиная с второй строки (предполагается, что первая строка - заголовки)
    for row_number, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        # Проверяете значение в столбце L (12-й столбец, так как индексация начинается с 1)
        if row[11] == "SL":
            # Если значение равно "SL", закрашиваете строку в красный цвет
            for col_idx, value in enumerate(row, start=1):
                cell = sheet.cell(row=row_number, column=col_idx)
                cell.fill = PatternFill(start_color="FF4040", end_color="FF4040", fill_type="solid")
                cell.border = Border(left=Side(style='thin', color='000000'),
                                    right=Side(style='thin', color='000000'),
                                    top=Side(style='thin', color='000000'),
                                    bottom=Side(style='thin', color='000000'))
        if row[11] == "TP":
            # Если значение равно "TL", закрашиваете строку в green цвет
            for col_idx, value in enumerate(row, start=1):
                cell = sheet.cell(row=row_number, column=col_idx)
                cell.fill = PatternFill(start_color="9FEE00", end_color="9FEE00", fill_type="solid")
                cell.border = Border(left=Side(style='thin', color='000000'),
                                    right=Side(style='thin', color='000000'),
                                    top=Side(style='thin', color='000000'),
                                    bottom=Side(style='thin', color='000000'))
        if row[11] == "EXP":
            # Если значение равно "EXP", закрашиваете строку в blue цвет
            for col_idx, value in enumerate(row, start=1):
                cell = sheet.cell(row=row_number, column=col_idx)
                cell.fill = PatternFill(start_color="66A3D2", end_color="66A3D2", fill_type="solid")
                cell.border = Border(left=Side(style='thin', color='000000'),
                                    right=Side(style='thin', color='000000'),
                                    top=Side(style='thin', color='000000'),
                                    bottom=Side(style='thin', color='000000'))
        if row[11] == "NOT":
            # Если значение равно "NOT", закрашиваете строку в желтый цвет
            for col_idx, value in enumerate(row, start=1):
                cell = sheet.cell(row=row_number, column=col_idx)
                cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        
                # Устанавливаете границы ячейки
                cell.border = Border(left=Side(style='thin', color='000000'),
                                    right=Side(style='thin', color='000000'),
                                    top=Side(style='thin', color='000000'),
                                    bottom=Side(style='thin', color='000000'))
        
    # Сохраняете изменения
    workbook.save(FILE_NAME)
    workbook.close()

# print(load_data("EUR%2FCHF.E.FX",'24.01.2024 14:56:11',  '26.01.2024 00:00:00'))
# print(processing_excel_data(pd.read_excel('25-01-2024 - ToBot (1).xlsx')))
# manager = Process_manager('25-01-2024 - ToBot (1).xlsx')
# manager.main_data_processing()
# data = load_data('EUR%2FCHF.E.FX','25.01.2024 10:14:31',  '26.01.2024 09:00:00')

# print(Strategy_Tester(data, 'Sell','Stop', 0.9407, 0.9383, 0.9323))


# print(data)