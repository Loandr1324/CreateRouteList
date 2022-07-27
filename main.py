#Author Loik Andrey 7034@balancedv.ru

import pandas as pd
import os
from datetime import datetime
import random
import numpy as np
import sys

ROUTE_PATH = 'маршруты'
FOLDER = 'Исходные данные'

def serch_file(end_file):
    for item in os.listdir(FOLDER):
        if item.endswith(end_file):
            file = FOLDER + "/" + item
    return file

def date_schedule(df):
    """ Определяем дату начала и конца графика работы """
    mask_date_tabel = df == 'Отчетный период'
    df_date_row = df[mask_date_tabel].dropna(axis=0, how='all').index.values
    df_date_col = df[mask_date_tabel].dropna(axis=1, how='all').columns
    df_date_col_idx = df.columns.get_loc(df_date_col[0])
    date_start = datetime.strptime(df.iloc[df_date_row[0] + 2, df_date_col_idx], '%d.%m.%Y')
    date_end = datetime.strptime(df.iloc[df_date_row[0] + 2, df_date_col_idx + 4], '%d.%m.%Y')
    return date_end

def df_transformation(df):
    """
    Преобразование эксель файла в DataFrame для дальнеёшей работы
    :param DataFrame фактического графика работы из 1С ЗУП
    :return: DataFrame: индексы - табльные номера, имена колонок - дни месяца,
    значения - Я, В, ОД и т.п. условные обозначения явки в графике работы
    """

    mask_start = df == 'Номер \nпо \nпоряд- \nку'
    df_start_row = df[mask_start].dropna(axis=0, how='all').index.values
    df_start_col = df[mask_start].dropna(axis=1, how='all').columns
    df_start_col_idx = df.columns.get_loc(df_start_col[0])
    mask_end = df == 'Ответственное\nлицо '
    df_end_row = df[mask_end].dropna(axis=0, how='all').index.values

    df1 = df.iloc[df_start_row[0] + 7:df_start_row[1], df_start_col_idx + 3:df_start_col_idx + 37]
    df2 = df.iloc[df_start_row[1] + 7:df_end_row[0] - 1, df_start_col_idx + 3:df_start_col_idx + 37]

    df3 = pd.concat([df1, df2], ignore_index=True)
    df3 = df3.dropna(axis=0, how='all')
    df3 = df3.dropna(axis=1, how='all')
    df_len = len(df3)

    for i in range(1, df_len, 2):
        df3 = df3.drop(i)
    df3.reset_index(drop=True, inplace=True)

    df3_len = len(df3)
    df4 = pd.DataFrame()
    df5 = pd.DataFrame()
    for i in range(1, df3_len, 2):
        df3.iloc[i:i + 1, 0:2] = df3.iloc[i - 1:i, 0:2]
        df4 = pd.concat([df4, df3.iloc[i - 1:i, :16]], axis=0, ignore_index=False)
        df5 = pd.concat([df5, df3.iloc[i:i + 1, :]], axis=0, ignore_index=False)

    columns4 = ['Табельный номер', 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
    columns5 = ['Табельный номер', 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]
    df4.columns = columns4
    df5.columns = columns5
    df4.set_index(['Табельный номер'], inplace=True)
    df5.set_index(['Табельный номер'], inplace=True)
    df_out = pd.concat([df4, df5], axis=1, ignore_index=False, levels=['Табельный номер'])
    return df_out

def random_route(df):
    """
    Подставляем номера маршуртов в дни явки сотрудника
    :param DataFrame фактического графика работы из 1С ЗУП
    :return: DataFrame: индексы - табльные номера, имена колонок - дни месяца,
    значения - в дни Я - случайные номера маршрутов, в остальные дни подставляется NA
    """

    mask_day = df == 'Я' # Создаём маску по дням явки
    shape_df_driver = df.shape

    count_route = fcount(ROUTE_PATH) # Определяем количество файлов+1 в папке ROUTE_PATH

    # Создаём массив со случайными числами от 1 до count_route таким же размером как и DataFrame df
    random_array = np.random.randint(low=1, high=count_route, size=(shape_df_driver[0], shape_df_driver[1]))

    # Преобразовываем массив со случайными номерами маршрутов в DateFrame
    df_drv_rnd = pd.DataFrame(random_array, columns=df.columns, index=df.index)

    # Оставляем случайные номера маршрутов только в дни явки
    df[~mask_day] = None
    df[mask_day] = df_drv_rnd[mask_day]
    df = df.astype('Int64')
    return df

def fcount(path):
    """ Определяем количество файлов в папке """
    count = 0
    for f in os.listdir(path):
        if os.path.isfile(os.path.join(path, f)):
            count += 1
    return count+1

def ex_driver_lic(date_control):
    """
    Получаем данные по документам водителя и проверяем срок действия
    :return: DataFrame с данными по документам водителей
    """
    file_driv_license = FOLDER + '/водители.xlsx'

    df_driv_license = pd.read_excel(file_driv_license, index_col='Табельный номер')
    df_driv_license['Срок действия'] = pd.to_datetime(df_driv_license['Срок действия'], format='%d.%m.%Y')
    #date_control = datetime.strptime('31.05.2026', '%d.%m.%Y') # Для теста работы проверки срока действия
    mask = df_driv_license['Срок действия'] < date_control

    # Выводим сообщение и завершаем работу программы при нахождении закончившихся водительских удостоверений
    if len(df_driv_license[mask]) > 0:
        print(f"Срок действия прав истёк до {date_control.strftime('%d.%m.%Y')} у следующих водителей:")
        print(df_driv_license[mask])
        input(
            f'\nВведите данные новых водительских удостоверений в файле {file_driv_license} и запустите программу снова')
        sys.exit()

    return df_driv_license.drop('Срок действия', axis=1)

def read_cars():
    # Загрузка данных по автомобилям
    df = pd.read_excel(FOLDER + '/автомобили.xlsx', index_col='Табельный номер')
    # Создаём копию старых данных по автомобилям
    cars_to_excel(df, FOLDER + '/автомобили_до работы программы.xlsx')
    df['Расход'] = 0
    return df

def create_route(date_tabel, df_driver_route, df_drivers, df_cars):
    number_route = int(input('Введите номер первого путевого листа: '))

    for col_name, data in df_driver_route.items():
        data = data.dropna()
        if len(data) > 0:
            for index in data.index:
                #print("index:", index, "; data.loc[index]:", data.loc[index])
                #print(f'Номер путевого листа: {number_route}')

                # Определяем дату путевого листа
                if len(str(col_name)) == 1:
                    col_name = '0' +  str(col_name)
                date_route = str(col_name) + date_tabel.strftime(".%m.%Y")
                #print(f'Дата путевого листа: {date_route}')

                # Выбираем данные транспортного средства
                df_car = df_cars.loc[index]
                car = 'автомобиль ' + df_car['Автомобиль'] + ', государственный регистрационный знак ' + df_car[
                    'Номер автомобиля']
                print(f'Транспортное средство: {car}')

                # Опеределяем расход топлива исходя из сезона
                fuel_consumption = consumption_car(date_tabel, df_car)
                print(f'Расход топлива: {fuel_consumption}')

                fuel_type = df_car['Марка топлива']
                #print(f'Марка топлива: {fuel_type}')
                odometer = df_car['Показания одометра'] # Задаём показание одометра в начале маршрута
                print(f'Стартовые показания одометра для маршрута: {odometer}')

                # Выбираем данные водителя
                df_driver = df_drivers.loc[index]
                driver = df_driver['Водитель'] + ', водительское удостоверение № ' + df_driver['Права']
                print(f'Водитель: {driver}')

                # Загружаем путевой лист маршрута, рассчитываем расход топлива по маршруту
                print(f'Номер загружаемого маршрута: {data.loc[index]}')
                file_name_read = 'маршруты/' + str(data.loc[index]) + '.xlsx' # Наименование файла подставляемого маршрута
                # print(f'Файл шаблона для загрузки: {file_name_read}')
                route_driver = pd.read_excel(file_name_read, header=None)
                route_distance = route_driver.iloc[28, 2]
                print(f'Пробег по маршруту: {route_distance}')
                route_fuel_consumption = np.round( fuel_consumption * route_distance / 100, 2)
                print(f"Расход топлива по маршруту: {route_fuel_consumption}")

                # Подставляем данные в путевой лист
                route_driver.iloc[1, 3] = route_driver.iloc[1, 7]
                route_driver.iloc[1, 7] = None
                route_driver.iloc[5, 7] = number_route
                route_driver.iloc[5, 8] = date_route
                route_driver.iloc[9, 2] = car
                route_driver.iloc[11, 2] = fuel_consumption
                route_driver.iloc[13, 2] = fuel_type
                route_driver.iloc[15, 2] = driver
                route_driver.iloc[29, 2] = route_fuel_consumption
                route_driver.iloc[19, 6] = odometer

                odometer += route_distance
                route_driver.iloc[19, 8] = odometer

                number_route += 1
                print(f'Показания одометра после маршрута: {odometer}')
                df_cars.loc[index, 'Показания одометра'] = odometer # Записываем показание одометра в конце маршрута
                df_cars.loc[index, 'Расход'] += route_fuel_consumption


                file_name = create_folder(date_tabel, str(col_name), df_driver['Водитель'])
                route_to_excel(route_driver, file_name)
    #TODO закомментировать при тестах
    cars_to_excel(df_cars.drop(['Расход'], axis=1), FOLDER + '/автомобили.xlsx')
    return df_cars

def get_date(date):
    month_list = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
           'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']
    date_list = date.split('.')
    return (month_list[int(date_list[1]) - 1] + ' ' +
        date_list[2])

def create_df_all_cars(df):
    df_mod = df.drop(['Марка топлива', 'Расход топлива лето (04-10)', 'Расход топлива зима (11-03)'], axis=1)
    df_mod.rename(columns={'Показания одометра': 'Показания на начало'}, inplace=True)
    df_mod['Показания на конец'] = 0
    list = ['Автомобиль', 'Номер автомобиля', 'Показания на начало', 'Показания на конец','Расход']
    return df_mod[list]

def consumption_car(date, df):
    """
    Определяем расход топлива согласно сезона
    :param date: Дата в формате datetime
    :param df: DataFrame c одной строкой и колонками: 'Расход топлива лето (04-10)', 'Расход топлива зима (11-03)'
    :return: Расход топлива int
    """
    m_date_route = int(date.strftime("%m"))
    if m_date_route >= 4 and m_date_route <= 10:
        consumption = df['Расход топлива лето (04-10)']
    else:
        consumption = df['Расход топлива зима (11-03)']
    return consumption

def create_folder(date, day, driver):
    """
    Создаём директории по каждому водителю и месяцу и формируем наименование файла с полным путём к нему
    :param date: Дата в формате datetime
    :param driver: ФИО Водителя str
    :return: Полный путь и наименование самого файла str
    """
    folder = '_' + driver  # Определяем директорию водителя
    folder_month = folder + '/' + date.strftime("%Y.%m")
    # Создаём директорию по каждому водителю
    try:
        os.mkdir(folder)
    except:
        pass
    try:
        os.mkdir(folder_month)
    except:
        pass

    #print(f'Наименование папки: {folder}')
    #print(f"Наименование файла для записи: {folder_month + '/' + date.strftime("%Y.%m.") + str(col_name) + '.xlsx'}")
    return folder_month + '/' + date.strftime("%Y.%m.") + str(day) + '.xlsx'

def create_final_file(date_tabel, df_final):
    str_m_y_date = get_date(date_tabel.strftime(".%m.%Y"))  # Заменяем номер месяца на название месяца
    df_sample = pd.read_excel(FOLDER + '/Шаблон расчета фактического расхода топлива по машинам.xlsx', header=None)

    df_sample.iloc[7, 0] = str_m_y_date
    df_sample.iloc[13, 3] = np.round(df_final['Расход'].sum(), 0)
    i = 9
    for index in df_final.index:
        # Подставляем данные в итоговый файл
        df_sample.iloc[i, 0] = df_final.loc[index, 'Автомобиль'] + ' ' + df_final.loc[index, 'Номер автомобиля']
        df_sample.iloc[i, 1] = df_final.loc[index, 'Показания на начало']
        df_sample.iloc[i, 2] = df_final.loc[index, 'Показания на конец']
        df_sample.iloc[i, 3] = df_final.loc[index, 'Расход']
        i += 1
    file_name = date_tabel.strftime("%Y.%m ") + 'Расчет фактического расхода топлива по машинам.xlsx'
    final_to_excel(df_sample, file_name)
    return

def cars_to_excel(df, file_name):

    # Переносим индекс в колонку
    df['Табельный номер'] = df.index

    # Изменяем порядок колонок DataFrame
    list = ['Табельный номер', 'Автомобиль', 'Марка топлива', 'Расход топлива лето (04-10)',
            'Расход топлива зима (11-03)', 'Номер автомобиля', 'Показания одометра']
    df = df[list]

    # Сбрасываем встроенный формат заголовков pandas
    pd.io.formats.excel.ExcelFormatter.header_style = None

    # Открываем файл для записи
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    sheet_name = 'Sheet1' # Задаём имя вкладки
    workbook = writer.book # Открываем книгу для записи

    # Записываем данные на вкладку sheet_name
    df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)

    # Выбираем вкладку для форматирования
    wks1 = writer.sheets[sheet_name]

    # Задаём форматы для ячеек
    consumption_format = workbook.add_format({'align': 'center'})
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })
    header_format.set_text_wrap() # Устанавливаем перенос строки для формата

    # Изменяем ширину и формат колонок
    wks1.set_column('A:D', 12, None)
    wks1.set_column('B:B', 20, None)
    wks1.set_column('C:C', 15, None)
    wks1.set_column('D:E', 15, consumption_format)
    wks1.set_column('F:F', 12, None)
    wks1.set_column('G:G', 10, None)

    # Изменяем высоту строк
    wks1.set_row(0, 30, None)

    # Изменяем формат заголовка таблицы
    for col_num, value in enumerate(df.columns.values):
        wks1.write(0, col_num, value, header_format)

    # Записываем изменённый файл
    writer.save()

    return

def route_to_excel(df, file_name):
    # Открываем файл для записи
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    sheet_name = 'Sheet1'  # Задаём имя вкладки
    workbook = writer.book  # Открываем книгу для записи

    # Записываем данные на вкладку sheet_name
    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    # Выбираем вкладку для форматирования
    wks1 = writer.sheets[sheet_name]

    # Задаём форматы для ячеек
    consumption_format = workbook.add_format({'align': 'center'})
    header_row_format = workbook.add_format({
        'bold': True,
        'font_size': 9,
        'font_name': 'Arial',
        'valign': 'right'
    })
    header_route_format = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'font_name': 'Arial',
        'valign': 'right'
    })
    header_num_date_format = workbook.add_format({
        'bold': True,
        'font_size': 12,
        'font_name': 'Arial',
        'align': 'center'
    })
    header_list_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'text_wrap': True,
        'align': 'left'
    })
    header_cus_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'text_wrap': True
    })
    all_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial'
    })
    cus_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'align': 'center',
    })
    hader_table_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'text_wrap': True,
        'bold': True,
        'border': True
    })
    hader_table_row_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'text_wrap': True,
        'bold': True,
        'border': True,
        'align': 'center',
        'valign': 'vcenter'
    })

    table_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'border': True,
        'align': 'center',
        'valign': 'vcenter'
    })
    table_format2 = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'border': True,
        'align': 'left',
        'valign': 'vcenter'
    })

    time_start_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'border': True,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': 'hh:mm'
    })
    time_end_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'border': True,
        'align': 'center',
        'valign': 'vcenter',
        'num_format': 'hh:mm'
    })

    odometr_start_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'border': True,
        'align': 'center',
        'valign': 'top',
    })

    odometr_end_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'border': True,
        'align': 'center',
        'valign': 'bottom',
    })

    footer_cus_format = workbook.add_format({
        'bold': True,
        'font_size': 10,
        'font_name': 'Arial',
        'align': 'center'
    })
    footer_dir_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'bottom': True,
        'align': 'center',
    })
    footer_sign_format = workbook.add_format({
        'font_size': 8,
        'font_name': 'Arial',
        'align': 'center',
    })

    # Изменяем ширину и формат колонок
    wks1.set_column('A:A', 4.14, None)
    wks1.set_column('B:B', 11.86, None)
    wks1.set_column('C:C', 22.29, None)
    wks1.set_column('D:D', 12.57, None)
    wks1.set_column('E:E', 8.14, None)
    wks1.set_column('F:F', 4.43, None)
    wks1.set_column('G:G', 2.83, None)
    wks1.set_column('H:H', 5.0, None)
    wks1.set_column('I:I', 7.71, None)
    wks1.set_column('J:J', 10.14, None)

    # Изменяем формат строк
    for i in range(0, 3):
        wks1.set_row(i, None, header_row_format)

    wks1.set_row(5, None, header_num_date_format)
    for i in range(7, 12 , 2):
        wks1.set_row(i, 25.5, header_list_format)
    wks1.write(11, 2, df.iloc[11, 2], cus_format)
    for i in range(13, 36):
        wks1.set_row(i, None, all_format)
    wks1.set_row(18, 40.5, None)

    # Объединяем ячейки
    wks1.merge_range(0, 8, 0, 9, None, None)
    wks1.merge_range(1, 3, 1, 9, None, None)
    wks1.merge_range(2, 8, 2, 9, None, None)
    wks1.merge_range(5, 0, 5, 5, 'Путевой лист автомобиля', header_route_format)
    wks1.merge_range(5, 8, 5, 9, None, None)
    wks1.merge_range(7, 0, 7, 1, None, None)
    wks1.merge_range(7, 2, 7, 9, None, None)
    wks1.merge_range(9, 0, 9, 1, None, None)
    wks1.merge_range(9, 2, 9, 9, None, None)
    wks1.merge_range(11, 0, 11, 1, None, None)

    # Оформление заголовка таблицы
    wks1.merge_range(17, 0, 18, 0, df.iloc[17, 0], hader_table_row_format)
    wks1.merge_range(17, 1, 17, 5, df.iloc[17, 1], hader_table_row_format)
    wks1.merge_range(17, 6, 17, 8, df.iloc[17, 6], hader_table_row_format)
    wks1.merge_range(18, 4, 18, 5, df.iloc[18, 4], hader_table_format)
    wks1.merge_range(18, 6, 18, 7, df.iloc[18, 6], hader_table_format)
    wks1.merge_range(17, 9, 18, 9, df.iloc[17, 9], hader_table_row_format)
    wks1.write(18,1, df.iloc[18, 1],hader_table_format)
    wks1.write(18,2, df.iloc[18, 2],hader_table_format)
    wks1.write(18,3, df.iloc[18, 3],hader_table_format)
    wks1.write(18,8, df.iloc[18, 8],hader_table_format)

    # Оформляем таблицу
    wks1.merge_range(19, 6, 26, 7, df.iloc[19, 6], odometr_start_format)
    wks1.merge_range(19, 8, 26, 8, df.iloc[19, 8], odometr_end_format)
    wks1.merge_range(19, 9, 26, 9, None, table_format)
    for i in range(19, 27):
        wks1.set_row(i, 24, None)
        wks1.write(i, 0, is_nan(df.iloc[i, 0]), table_format)
        wks1.write(i, 1, is_nan(df.iloc[i, 1]), time_start_format)
        wks1.write(i, 2, is_nan(df.iloc[i, 2]), table_format2)
        wks1.write(i, 3, is_nan(df.iloc[i, 3]), table_format)
        wks1.merge_range(i, 4, i, 5, is_nan(df.iloc[i, 4]), time_end_format)

    # Заполняем подвал
    wks1.write(28, 2, df.iloc[28, 2], footer_cus_format)
    wks1.write(29, 2, df.iloc[29, 2], footer_cus_format)
    wks1.merge_range(32, 0, 32, 1, df.iloc[32, 0], footer_dir_format)
    wks1.merge_range(32, 3, 32, 4, None, footer_dir_format)
    wks1.merge_range(32, 5, 32, 9, df.iloc[32, 7], footer_dir_format)
    wks1.merge_range(33, 0, 33, 1, df.iloc[33, 0], footer_sign_format)
    wks1.merge_range(33, 3, 33, 4, df.iloc[33, 3], footer_sign_format)
    wks1.merge_range(33, 5, 33, 9, df.iloc[33, 7], footer_sign_format)

    # Задаём параметры печати
    wks1.set_margins(left=0.4, right=0.4, top=0.4, bottom=0.4)
    # Записываем изменённый файл
    writer.save()

def is_nan(x):
    if x is np.nan:
        x = ''
    return x

def final_to_excel(df, file_name):

    # Открываем файл для записи
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    sheet_name = 'Sheet1' # Задаём имя вкладки
    workbook = writer.book # Открываем книгу для записи

    df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)

    # Выбираем вкладку для форматирования
    wks1 = writer.sheets[sheet_name]

    # Изменяем ширину и формат колонок
    wks1.set_column('A:A', 26.29, None)
    wks1.set_column('B:B', 10.43, None)
    wks1.set_column('C:C', 10.71, None)
    wks1.set_column('D:D', 12, None)


    # Форматируем заголок
    header_format = workbook.add_format({
        'bold': True,
        'font_size': 9,
        'font_name': 'Arial',
        'align': 'right'
    })
    wks1.merge_range(0, 2, 0, 3, df.iloc[0, 2], header_format)
    wks1.merge_range(1, 0, 1, 3, df.iloc[1, 0], header_format)
    wks1.merge_range(2, 2, 2, 3, df.iloc[2, 2], header_format)

    header_row_format = workbook.add_format({
        'font_size': 10,
        'font_name': 'Arial',
        'text_wrap': True,
    })
    wks1.set_row(4, None, header_row_format)
    wks1.set_row(5, 30, header_row_format)
    wks1.merge_range(5, 0, 5, 3, None, None)

    header_title_format = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'font_name': 'Arial',
        'align': 'center',
        'valign': 'vcenter'
    })
    wks1.set_row(6, 30, None)
    wks1.merge_range(6, 0, 6, 3, df.iloc[6, 0], header_title_format)

    header_month_format = workbook.add_format({
        'italic': True,
        'font_size': 14,
        'font_name': 'Calibri',
        'align': 'center',
        'valign': 'vcenter'
    })
    wks1.set_row(7, 40.5, None)
    wks1.merge_range(7, 0, 7, 3, df.iloc[7, 0], header_month_format)

    # Форматируем таблицу
    table_header_format = workbook.add_format({
        'border': True,
        'bold': True,
        'italic': True,
        'font_color': 'red',
        'text_wrap': True,
        'valign': 'top'
    })
    wks1.set_row(8, 30, None)
    for i in range(0,4):
        wks1.write(8, i, is_nan(df.iloc[8, i]), table_header_format)

    table_col1_format = workbook.add_format({
        'border': True,
        'bold': True,
        'font_color': 'navy'
    })
    table_data_format = workbook.add_format({
        'border': True,
    })

    for i in range(9, 13):
        wks1.write(i, 0, is_nan(df.iloc[i, 0]), table_col1_format)
        wks1.write(i, 1, is_nan(df.iloc[i, 1]), table_data_format)
        wks1.write(i, 2, is_nan(df.iloc[i, 2]), table_data_format)
        wks1.write(i, 3, is_nan(df.iloc[i, 3]), table_data_format)

    table_footer_col1_format = workbook.add_format({
        'border': True,
        'bold': True,
        'font_color': 'navy',
        'font_size': 14,
    })
    table_footer_format = workbook.add_format({
        'border': True,
        'bold': True,
        'font_size': 12,
    })
    wks1.set_row(13, 18.75, None)
    wks1.write(13, 0, is_nan(df.iloc[13, 0]), table_footer_col1_format)
    for i in range(1,4):
        wks1.write(13, i, is_nan(df.iloc[13, i]), table_footer_format)

    # Записываем изменённый файл
    writer.save()
    return

def Run():
    # Находим файл с графиком работы
    file_schedule = serch_file('табель.xlsx')

    # Считываем график работы
    df_schedule = pd.read_excel(file_schedule)

    # Определяем конечную дату загруженного графика работы
    date_tabel_end = date_schedule(df_schedule)

    # Преобразовываем DataFrame с графиком работы для удобства дальнеёшей работы
    df_driver_schedule = df_transformation(df_schedule)

    # Подставляем случайные номера маршрутов в дни явки сотрудников
    df_driver_route = random_route(df_driver_schedule)

    # Загружаем файл с документами водителей и проверяем срок действия
    df_drivers = ex_driver_lic(date_tabel_end)

    # Загружаем файл с данными по автомобилям
    df_cars = read_cars()

    # Создаём DataFrame для итогового отчёта за месяц по всем автомобилям
    df_all_cars = create_df_all_cars(df_cars)

    # Заполняем и сохраняем все графики и возвращаем данные на конец месяца
    df_cars_after = create_route(date_tabel_end, df_driver_route, df_drivers, df_cars)

    # Подготавливаем DataFrame для записи в итоговый файл за месяц
    df_all_cars[['Показания на конец', 'Расход']] = df_cars_after[['Показания одометра', 'Расход']]

    # Создаём итоговый файл
    create_final_file(date_tabel_end, df_all_cars)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    Run()
    input('Программа закончила работу.\nНажмите Enter для выхода')
