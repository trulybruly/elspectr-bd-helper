#!./venv/bin/python3.9
import argparse
import pandas as pd
import xlrd
import numpy as np
from transliterate import slugify
from datetime import datetime
import time
import multiprocessing as mp

parser = argparse.ArgumentParser(usage="elspectr", description="4 xls files into csv for e-shop.")
parser.add_argument("tables", help="excell files", nargs="+")
parser.add_argument("-o", "--output", required=True, help="output location (required)")
# parser.add_argument("-e","third",required=True, help="Third file")
# parser.add_argument("-r","fourth",required=True, help="Fourth file")
args = parser.parse_args()


# def shopChecker (df):
#    str = df['Unnamed: 1'][3]
#    if str.lower.find('революци') != -1:
#        return 'varya'
#    if str.lower.find('свобод') != -1:
#        return 'varya'
#    if str.lower.find('революци') != -1:
#        return 'varya'
#    if str.lower.find('революци') != -1:
#        return 'varya'

def getBGColor(book, bSheet, bRow, col):
    xfx = bSheet.cell_xf_index(bRow, col)
    xf = book.xf_list[xfx]
    bgx = xf.background.pattern_colour_index
    pattern_colour = book.colour_map[bgx]
    # Actually, despite the name, the background colour is not the background colour.
    # background_colour_index = xf.background.background_colour_index
    # background_colour = book.colour_map[background_colour_index]
    return pattern_colour


def groupChecker(excel_book, row):
    black = (0, 0, 0)
    grey = (128, 128, 128)
    l_grey = (160, 160, 164)
    su_l_grey = (192, 192, 192)

    color = getBGColor(excel_book, excel_book[0], row, 1)
    if color == black:
        return 0
    if color == grey:
        return 1
    if color == l_grey:
        return 2
    if color == su_l_grey:
        return 3
    return 4


def testTableMaker():
    table = pd.DataFrame({'ID': [1, 2], 'UI': ['tg', 'yu']})
    table = table.append(pd.Series(), ignore_index=True)
    return table


def tableMaker():
    return pd.DataFrame(
        {'ID': [], 'Тип': [], 'Артикул': [], 'Имя': [], 'Опубликован': [], 'рекомендуемый?': [],
         'Видимость в каталоге': [], 'Краткое описание': [], 'Описание': [],
         'Дата начала действия продажной цены': [],
         'Дата окончания действия продажной цены': [], 'Статус налога': [], 'Налоговый класс': [],
         'В наличии?': [], 'Продано индивидуально?': [], 'Вес (kg)': [],
         'Разрешить отзывы от клиентов?': [], 'Примечание к покупке': [], 'Цена распродажи': [],
         'Базовая цена': [], 'Категории': [], 'Метки': [], 'Класс доставки': [], 'Изображения': [],
         'Родительский': [], 'Сгруппированные товары': [], 'Апсейл': [], 'Кросселы': [],
         'Внешний URL': [], 'Текст кнопки': [], 'Позиция': [], 'Имя атрибута 1': [],
         'Значение(-я) атрибута(-ов) 1': [],
         'Видимость атрибута 1': [], 'Глобальный атрибут 1': []}, dtype='string'
    )


def isThereSmthSimilar(name, finalTable, row_count_divided_by_5_plus_1):
    # for i in range((finalTable.isna()['ID'].sum() + finalTable.notna()['ID'].sum())//5):
    for i in range(row_count_divided_by_5_plus_1 - 1):
        # if finalTable.notna()['Имя'][i * 5]:
        #     if finalTable['Имя'][i * 5] == name:
        #         return i * 5
        if finalTable['Имя'][i * 5] == name:
            return i * 5
    return -1


def makePictureLink(name):
    link = 'https://new.elspectr.ru/wp-content/uploads/' + slugify(name, language_code='ru') + '.jpg'
    return link


def tableFiller(finalTable: pd.DataFrame, id_: int, name: str, price: str, sale_price: str, group: str, subgroup: str,
                sub_sub_group: str, sub_sub_sub_group: str, shop: str, global_product_counter: int, short_expl:str):
    # attributeName = 'Магазин'
    # finalTable = finalTable.copy()

    similarity = isThereSmthSimilar(name, finalTable, global_product_counter)
    if similarity != -1:
        for k in range(2,5):
            if finalTable.isna()['Имя'][similarity+ k]:
                finalTable['ID'][similarity + k] = str(similarity + k)
                finalTable['Тип'][similarity + k] = 'variation'
                # finalTable['Артикул'][similarity+1+i] = SKU
                finalTable['Имя'][similarity + k] = name + ' - ' + shop
                finalTable['Опубликован'][similarity + k] = '1'
                finalTable['рекомендуемый?'][similarity + k] = '0'
                finalTable['Видимость в каталоге'][similarity + k] = 'visible'
                # finalTable['Краткое описание'][similarity + k] = short_expl
                finalTable['Статус налога'][similarity + k] = 'taxable'
                finalTable['Налоговый класс'][similarity + k] = 'parent'
                finalTable['В наличии?'][similarity + k] = '1'
                # finalTable['Возможен ли предзаказ?'][similarity + 1 + k] = 0
                finalTable['Продано индивидуально?'][similarity + k] = '0'
                finalTable['Разрешить отзывы от клиентов?'][similarity + k] = '0'
                finalTable['Класс доставки'][similarity + k] = shop
                if sale_price != '-1':
                    finalTable['Цена распродажи'][similarity + k] = sale_price
                finalTable['Базовая цена'][similarity + k] = price
                # finalTable['Категории'][similarity+1+i] = categories
                # finalTable['Изображения'][similarity+1+i] = picture + ', ' + picture
                finalTable['Класс доставки'][similarity + k] = shop
                finalTable['Родительский'][similarity + k] = finalTable['Артикул'][similarity]
                finalTable['Позиция'][similarity + k] = k
                finalTable['Имя атрибута 1'][similarity + k] = 'Магазин'
                finalTable['Значение(-я) атрибута(-ов) 1'][similarity + k] = shop
                # finalTable['Видимость атрибута 1'][similarity+1+i] = 1
                finalTable['Глобальный атрибут 1'][similarity + k] = '1'

                return finalTable, global_product_counter

    else:
        lastIndex = len(finalTable['ID'])

        # filling 5 rows by NaN
        # dtypes = [
        #     'int', 'str', 'int', 'str', 'int', 'int', 'str', 'str', 'str', 'str', 'str', 'str',# last - Статус налога
        #     'str', 'int', 'int', 'float', 'int', 'str', 'float', 'float', 'str', 'str', 'str', # last - Класс доставки
        #     'str', 'int', 'str', 'str', 'str', 'str', 'str', 'int', 'str', 'str', 'int', 'int'
        #     ]
        for i in range(5):
            finalTable = finalTable.append(pd.Series(dtype='string'), ignore_index=True)
        global_product_counter += 1

        SKU = lastIndex + 10000
        categories = group
        if subgroup != '':
            categories += ', ' + group + ' > ' + subgroup
            if sub_sub_group != '':
                categories += ', ' + group + ' > ' + subgroup + ' > ' + sub_sub_group
                if sub_sub_sub_group != '':
                    categories += ', ' + group + ' > ' + subgroup + ' > ' + sub_sub_group + ' > ' + sub_sub_sub_group
        picture = makePictureLink(name)

        # parent product
        finalTable['ID'][lastIndex] = str(lastIndex + 1)
        finalTable['Тип'][lastIndex] = 'variable'
        finalTable['Артикул'][lastIndex] = str(SKU)
        finalTable['Имя'][lastIndex] = str(name)
        finalTable['Опубликован'][lastIndex] = '1'
        finalTable['рекомендуемый?'][lastIndex] = '0'
        finalTable['Видимость в каталоге'][lastIndex] = 'visible'
        finalTable['Статус налога'][lastIndex] = 'taxable'
        finalTable['Краткое описание'][lastIndex] = 'Цена указана за ' + short_expl
        finalTable['В наличии?'][lastIndex] = '1'
        # finalTable['Возможен ли предзаказ?'][lastIndex] = 0
        finalTable['Продано индивидуально?'][lastIndex] = '0'
        finalTable['Разрешить отзывы от клиентов?'][lastIndex] = '1'
        finalTable['Категории'][lastIndex] = categories
        finalTable['Изображения'][lastIndex] = picture + ', ' + picture
        finalTable['Позиция'][lastIndex] = '0'
        finalTable['Имя атрибута 1'][lastIndex] = 'Магазин'
        finalTable['Значение(-я) атрибута(-ов) 1'][
            lastIndex] = 'ул. Плотникова 4, ул. Свободы 67, ул. Страж Революции 4, ул. Телеграфная 53'
        finalTable['Видимость атрибута 1'][lastIndex] = '1'
        finalTable['Глобальный атрибут 1'][lastIndex] = '1'

        # first child product
        finalTable['ID'][lastIndex + 1] = str(lastIndex + 2)
        finalTable['Тип'][lastIndex + 1] = 'variation'
        # finalTable['Артикул'][lastIndex+1] = SKU
        finalTable['Имя'][lastIndex + 1] = name + ' - ' + shop
        finalTable['Опубликован'][lastIndex + 1] = '1'
        finalTable['рекомендуемый?'][lastIndex + 1] = '0'
        finalTable['Видимость в каталоге'][lastIndex + 1] = 'visible'
        finalTable['Статус налога'][lastIndex + 1] = 'taxable'
        # finalTable['Краткое описание'][lastIndex] = short_expl
        finalTable['Налоговый класс'][lastIndex + 1] = 'parent'
        finalTable['В наличии?'][lastIndex + 1] = '1'
        # finalTable['Возможен ли предзаказ?'][lastIndex + 1] = 0
        finalTable['Продано индивидуально?'][lastIndex + 1] = '0'
        finalTable['Разрешить отзывы от клиентов?'][lastIndex + 1] = '0'
        finalTable['Класс доставки'][lastIndex + 1] = shop
        if sale_price != '-1':
            finalTable['Цена распродажи'][lastIndex + 1] = str(sale_price)
        finalTable['Базовая цена'][lastIndex + 1] = str(price)
        # finalTable['Категории'][lastIndex+1] = categories
        # finalTable['Изображения'][lastIndex+1] = picture + ', ' + picture
        finalTable['Класс доставки'][lastIndex + 1] = shop
        finalTable['Родительский'][lastIndex + 1] = finalTable['Артикул'][lastIndex]
        finalTable['Позиция'][lastIndex + 1] = '1'
        finalTable['Имя атрибута 1'][lastIndex + 1] = 'Магазин'
        finalTable['Значение(-я) атрибута(-ов) 1'][lastIndex + 1] = shop
        # finalTable['Видимость атрибута 1'][similarity+1+i] = 1
        finalTable['Глобальный атрибут 1'][lastIndex + 1] = '1'

        return finalTable, global_product_counter


if __name__ == '__main__':
    pd.options.mode.chained_assignment = None
    start_time = time.time()
    print('---------')
    print(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    # myTable=pd.DataFrame()
    myTable = tableMaker()
    # print(f'fisrt {type(myTable)}')
    _id = 1
    how_many_global_products = 0
    # myTable = myTable.dropna(how='all')

    atribut = ('ул. Плотникова 4', 'ул. Свободы 67', 'ул. Страж Революции 4', 'ул. Телеграфная 53')

    for i in range(len(args.tables)):
        print('---\nProcessing: '+atribut[i])
        table_time = time.time()

        shop = atribut[i]
        df = pd.DataFrame(pd.read_excel(args.tables[i]))
        tableSize = len(df)
        del df

        xlsBook = xlrd.open_workbook(args.tables[i], formatting_info=True)
        sheet = xlsBook[0]
        # print(df)
        # reading rows

        group = ''
        sub_group = ''
        subx2_group = ''
        subx3_group = ''

        for row in range(3000, tableSize):
            if sheet[row][1].value==' ' or sheet[row][1].value=='':
                continue

            # check group
            color = groupChecker(xlsBook, row)
            if color == 0:
                group = sheet[row][1].value
                sub_group = ''
                subx2_group = ''
                subx3_group = ''
            if color == 1:
                sub_group = sheet[row][1].value
                subx2_group = ''
                subx3_group = ''
            if color == 2:
                subx2_group = sheet[row][1].value
                subx3_group = ''
            if color == 3:
                subx3_group = sheet[row][1].value
            if color == 4:
                # myTable = pd.DataFrame(
                #     {'ID': [], 'Тип': [], 'Артикул': [], 'Имя': [], 'Опубликован': [], 'рекомендуемый?': [],
                #      'Видимость в каталоге': [], 'Краткое описание': [], 'Описание': [],
                #      'Дата начала действия продажной цены': [],
                #      'Дата окончания действия продажной цены': [], 'Статус налога': [], 'Налоговый класс': [],
                #      'В наличии?': [], 'Продано индивидуально?': [], 'Вес (kg)': [],
                #      'Разрешить отзывы от клиентов?': [], 'Примечание к покупке': [], 'Цена распродажи': [],
                #      'Базовая цена': [], 'Категории': [], 'Метки': [], 'Класс доставки': [], 'Изображения': [],
                #      'Родительский': [], 'Сгруппированные товары': [], 'Апсейл': [], 'Кросселы': [],
                #      'Внешний URL': [], 'Текст кнопки': [], 'Позиция': [], 'Имя атрибута 1': [],
                #      'Значение(-я) атрибута(-ов) 1': [],
                #      'Видимость атрибута 1': [], 'Глобальный атрибут 1': []}
                # )
                # print(f'second {type(myTable)}')
                # print(row)
                myTable, how_many_global_products = tableFiller(myTable.copy(), _id, sheet[row][1].value,
                                                                sheet[row][2].value, '-1', group, sub_group,
                                                                subx2_group, subx3_group, shop,
                                                                how_many_global_products, sheet[row][3].value)
                _id += 1

        del xlsBook
        del sheet

        print('processing time: '+str(time.time()-table_time)+', s')

    # deleting empty rows
    myTable = myTable.dropna(how='all')

    time_for_naming = datetime.now()
    fileName = args.output + '/prices_woo_' + time_for_naming.strftime('%Y-%m-%d_%H-%M-%S') + '.csv'
    myTable.to_csv(fileName, index=False)
    print(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    print('---------')
    print('execution_time')
    print(time.time() - start_time)
