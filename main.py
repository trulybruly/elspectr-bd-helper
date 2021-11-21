#!./venv/bin/python3.9
import argparse
import pandas as pd
import xlrd
import numpy as np
from transliterate import slugify
from datetime import datetime
import time

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

    color = getBGColor(excel_book, excel_book[0], row, 1)
    if color == black:
        return 0
    if color == grey:
        return 1
    return 2


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
         'Видимость атрибута 1': [], 'Глобальный атрибут 1': []}
    )


def isThereSmthSimilar(name, finalTable):
    for i in range(finalTable.isna()['ID'].sum() + finalTable.notna()['ID'].sum()):
        if finalTable.notna()['Имя'][i]:
            if finalTable['Имя'][i] == name:
                return i
    return -1


def makePictureLink(name):
    link = 'https://new.elspectr.ru/wp-content/uploads/' + slugify(name, language_code='ru') + '.jpg'
    return link


def tableFiller(finalTable: pd.DataFrame, id, name, price, sale_price, group, subgroup, shop):
    # attributeName = 'Магазин'
    # finalTable = finalTable.copy()

    similarity = isThereSmthSimilar(name, finalTable)
    if similarity != -1:
        for k in range(3):
            if finalTable.isna()['Имя'][similarity + 1 + k]:
                finalTable['ID'][similarity + 1 + k] = similarity + 1 + k
                finalTable['Тип'][similarity + 1 + k] = 'variation'
                # finalTable['Артикул'][similarity+1+i] = SKU
                finalTable['Имя'][similarity + 1 + k] = name + ' - ' + shop
                finalTable['Опубликован'][similarity + 1 + k] = 1
                finalTable['рекомендуемый?'][similarity + 1 + k] = 0
                finalTable['Видимость в каталоге'][similarity + 1 + k] = 'visible'
                finalTable['Статус налога'][similarity + 1 + k] = 'taxable'
                finalTable['Налоговый класс'][similarity + 1 + k] = 'parent'
                finalTable['В наличии?'][similarity + 1 + k] = 1
                # finalTable['Возможен ли предзаказ?'][similarity + 1 + k] = 0
                finalTable['Продано индивидуально?'][similarity + 1 + k] = 0
                finalTable['Разрешить отзывы от клиентов?'][similarity + 1 + k] = 0
                finalTable['Класс доставки'][similarity + 1 + k] = shop
                finalTable['Цена распродажи'][similarity + 1 + k] = sale_price
                finalTable['Базовая цена'][similarity + 1 + k] = price
                # finalTable['Категории'][similarity+1+i] = categories
                # finalTable['Изображения'][similarity+1+i] = picture + ', ' + picture
                finalTable['Класс доставки'][similarity + 1 + k] = shop
                finalTable['Родительский'][similarity + 1 + k] = finalTable['Артикул'][similarity]
                finalTable['Позиция'][similarity + 1 + k] = k + 1
                finalTable['Имя атрибута 1'][similarity + 1 + k] = 'Магазин'
                finalTable['Значение(-я) атрибута(-ов) 1'][similarity + 1 + k] = shop
                # finalTable['Видимость атрибута 1'][similarity+1+i] = 1
                finalTable['Глобальный атрибут 1'][similarity + 1 + k] = 1

                return finalTable

    else:
        lastIndex = len(finalTable['ID']) + 1

        # filling 5 rows by NaN
        for i in range(5):
            finalTable = finalTable.append(pd.Series(dtype='object'), ignore_index=True)

        SKU = lastIndex + 10000
        categories = group
        if subgroup != '':
            categories = categories + ', ' + group + ' > ' + subgroup
        picture = makePictureLink(name)

        # parent product
        finalTable['ID'][lastIndex] = id
        finalTable['Тип'][lastIndex] = 'variable'
        finalTable['Артикул'][lastIndex] = SKU
        finalTable['Имя'][lastIndex] = name
        finalTable['Опубликован'][lastIndex] = 1
        finalTable['рекомендуемый?'][lastIndex] = 0
        finalTable['Видимость в каталоге'][lastIndex] = 'visible'
        finalTable['Статус налога'][lastIndex] = 'taxable'
        finalTable['В наличии?'][lastIndex] = 1
        # finalTable['Возможен ли предзаказ?'][lastIndex] = 0
        finalTable['Продано индивидуально?'][lastIndex] = 0
        finalTable['Разрешить отзывы от клиентов?'][lastIndex] = 1
        finalTable['Категории'][lastIndex] = categories
        finalTable['Изображения'][lastIndex] = picture + ', ' + picture
        finalTable['Позиция'][lastIndex] = 0
        finalTable['Имя атрибута 1'][lastIndex] = 'Магазин'
        finalTable['Значение(-я) атрибута(-ов) 1'][
            lastIndex] = 'ул. Плотникова 4, ул. Свободы 67, ул. Страж Революции 4, ул. Телеграфная 53'
        finalTable['Видимость атрибута 1'][lastIndex] = 1
        finalTable['Глобальный атрибут 1'][lastIndex] = 1

        # first child product
        finalTable['ID'][lastIndex + 1] = lastIndex + 1
        finalTable['Тип'][lastIndex + 1] = 'variation'
        # finalTable['Артикул'][lastIndex+1] = SKU
        finalTable['Имя'][lastIndex + 1] = name + ' - ' + shop
        finalTable['Опубликован'][lastIndex + 1] = 1
        finalTable['рекомендуемый?'][lastIndex + 1] = 0
        finalTable['Видимость в каталоге'][lastIndex + 1] = 'visible'
        finalTable['Статус налога'][lastIndex + 1] = 'taxable'
        finalTable['Налоговый класс'][lastIndex + 1] = 'parent'
        finalTable['В наличии?'][lastIndex + 1] = 1
        # finalTable['Возможен ли предзаказ?'][lastIndex + 1] = 0
        finalTable['Продано индивидуально?'][lastIndex + 1] = 0
        finalTable['Разрешить отзывы от клиентов?'][lastIndex + 1] = 0
        finalTable['Класс доставки'][lastIndex + 1] = shop
        if sale_price != -1:
            finalTable['Цена распродажи'][lastIndex + 1] = sale_price
        finalTable['Базовая цена'][lastIndex + 1] = price
        # finalTable['Категории'][lastIndex+1] = categories
        # finalTable['Изображения'][lastIndex+1] = picture + ', ' + picture
        finalTable['Класс доставки'][lastIndex + 1] = shop
        finalTable['Родительский'][lastIndex + 1] = finalTable['Артикул'][lastIndex]
        finalTable['Позиция'][lastIndex + 1] = 1
        finalTable['Имя атрибута 1'][lastIndex + 1] = 'Магазин'
        finalTable['Значение(-я) атрибута(-ов) 1'][lastIndex + 1] = shop
        # finalTable['Видимость атрибута 1'][similarity+1+i] = 1
        finalTable['Глобальный атрибут 1'][lastIndex + 1] = 1

        return finalTable


if __name__ == '__main__':
    start_time=time.time()
    print('---------')
    print(datetime.now())
    # myTable=pd.DataFrame()
    myTable = tableMaker()
    print(f'fisrt {type(myTable)}')
    _id = 0
    # myTable = myTable.dropna(how='all')

    atribut = ('ул. Плотникова 4', 'ул. Свободы 67', 'ул. Страж Революции 4', 'ул. Телеграфная 53')

    for i in range(len(args.tables)):
        print(i)
        shop = atribut[i]
        df = pd.DataFrame(pd.read_excel(args.tables[i]))
        tableSize = len(df)
        del df

        xlsBook = xlrd.open_workbook(args.tables[i], formatting_info=True)
        sheet = xlsBook[0]
        # print(df)
        # reading rows

        row = 11
        for row in range(11, tableSize):
            # check group
            group = ''
            subgroup = ''
            color = groupChecker(xlsBook, row)
            if color == 0:
                group = sheet[row][1].value
                subgroup = ''
            if color == 1:
                subgroup = sheet[row][1].value
            if color == 2:
                price = sheet[row][2]
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
                myTable = tableFiller(myTable.copy(), _id, sheet[row][1].value, price, -1, group, subgroup, shop)

    # deleting empty rows
    myTable = myTable.dropna(how='all')

    time = datetime.now()
    fileName = args.output + '/prices_woo_' + time.strftime('%Y-%m-%d_%H-%M-%S') + '.csv'
    myTable.to_csv(fileName, index=False)
    print(datetime.now())
    print('---------')
    print('execution_time')
    print(time.time()-start_time)
