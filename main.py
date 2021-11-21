#import csv
import argparse
import pandas as pd
import xlrd
from transliterate import slugify
from datetime import datetime

parser = argparse.ArgumentParser(usage="elspectr", description="4 xls files into csv for shop. You have to write tables in specific order:\n avtozavod, svoboda, varya, telegraphnaya")
parser.add_argument("tables", help="First file", nargs="+")
#parser.add_argument("-w","second",required=True, help="Second file")
#parser.add_argument("-e","third",required=True, help="Third file")
#parser.add_argument("-r","fourth",required=True, help="Fourth file")
args = parser.parse_args()

#def shopChecker (df):
#    str = df['Unnamed: 1'][3]
#    if str.lower.find('революци') != -1:
#        return 'varya'
#    if str.lower.find('свобод') != -1:
#        return 'varya'
#    if str.lower.find('революци') != -1:
#        return 'varya'
#    if str.lower.find('революци') != -1:
#        return 'varya'

def getBGColor(book, sheet, row, col):
    xfx = sheet.cell_xf_index(row, col)
    xf = book.xf_list[xfx]
    bgx = xf.background.pattern_colour_index
    pattern_colour = book.colour_map[bgx]
    #Actually, despite the name, the background colour is not the background colour.
    #background_colour_index = xf.background.background_colour_index
    #background_colour = book.colour_map[background_colour_index]
    return pattern_colour



def groupChecker(xlsBook,row):
    black=(0,0,0)
    grey=(128,128,128)

    color=getBGColor(xlsBook,0,row,1)
    if color==black:
        return 0
    if color==grey:
        return 1
    return 2

def tableMaker():
    myTable = pd.DataFrame({'ID':[],'Тип':[],'Артикул':[],'Имя':[],'Опубликован':[],'рекомендуемый?':[],'Видимость в каталоге':[],'Краткое описание':[],'Описание':[],'Дата начала действия продажной цены':[],
                            'Дата окончания действия продажной цены':[],'Статус налога':[],'Налоговый класс':[],'В наличии?':[],'Продано индивидуально?':[],'Вес (kg)':[],
                            'Разрешить отзывы от клиентов?':[],'Примечание к покупке':[],'Цена распродажи':[],'Базовая цена':[],'Категории':[],'Метки':[],'Класс доставки':[],'Изображения':[],
                            'Родительский':[],'Сгруппированные товары':[],'Апсейл':[],'Кросселы':[],'Внешний URL':[],'Текст кнопки':[],'Позиция':[],'Имя атрибута 1':[],'Значение(-я) атрибута(-ов) 1':[],
                            'Видимость атрибута 1':[],'Глобальный атрибут 1':[]})
    return myTable

def isThereSmthSimilar(name, finalTable):
    for i in range(len(finalTable['ID'])):
        if finalTable['Имя'][i] == name:
            return i
    return -1


def makePictureLink(name):
    link = 'https://new.elspectr.ru/wp-content/uploads/' + slugify(name, language_code='ru')+'.jpg'
    return link

def tableFiller(finalTable:pd.DataFrame,id,name,price,sale_price,group,subgroup,shop):
    attributeName='Магазин'


    similarity = isThereSmthSimilar(name, finalTable)
    if similarity!=-1:
        for i in range(3):
            if finalTable.isna()['Имя'][similarity + 1 + i]:

                finalTable['ID'][similarity+1+i] = similarity+1+i
                finalTable['Тип'][similarity+1+i] = 'variation'
                #finalTable['Артикул'][similarity+1+i] = SKU
                finalTable['Имя'][similarity+1+i] = name+' - '+shop
                finalTable['Опубликован'][similarity+1+i] = 1
                finalTable['рекомендуемый?'][similarity+1+i] = 0
                finalTable['Видимость в каталоге'][similarity+1+i] = 'visible'
                finalTable['Статус налога'][similarity+1+i] = 'taxable'
                finalTable['Налоговый класс'][similarity+1+i] = 'parent'
                finalTable['В наличии?'][similarity+1+i] = 1
                finalTable['Возможен ли предзаказ?'][similarity+1+i] = 0
                finalTable['Продано индивидуально?'][similarity+1+i] = 0
                finalTable['Разрешить отзывы от клиентов?'][similarity+1+i] = 0
                finalTable['Класс доставки'][similarity+1+i] = shop
                finalTable['Цена распродажи'][similarity+1+i] = sale_price
                finalTable['Базовая цена'][similarity+1+i] = price
                #finalTable['Категории'][similarity+1+i] = categories
                #finalTable['Изображения'][similarity+1+i] = picture + ', ' + picture
                finalTable['Класс доставки'][similarity+1+i] = shop
                finalTable['Родительский'][similarity+1+i] = finalTable['Артикул'][similarity]
                finalTable['Позиция'][similarity+1+i] = i+1
                finalTable['Имя атрибута 1'][similarity+1+i] = 'Магазин'
                finalTable['Значение(-я) атрибута(-ов) 1'][similarity+1+i] = shop
                #finalTable['Видимость атрибута 1'][similarity+1+i] = 1
                finalTable['Глобальный атрибут 1'][similarity+1+i] = 1

                return

    else:
        lastIndex=len(finalTable['ID'])

        #filling 5 rows by NaN
        for i in range(5):
            finalTable = finalTable.append(pd.Series(),ignore_index=True)

        SKU=lastIndex+10000
        categories = group
        if subgroup!='':
            categories = categories + ', ' + group + ' > ' + subgroup
        picture = makePictureLink(name)

        #parent product
        finalTable['ID'][lastIndex] = id
        finalTable['Тип'][lastIndex] = 'variable'
        finalTable['Артикул'][lastIndex] = SKU
        finalTable['Имя'][lastIndex] = name
        finalTable['Опубликован'][lastIndex] = 1
        finalTable['рекомендуемый?'][lastIndex] = 0
        finalTable['Видимость в каталоге'][lastIndex] = 'visible'
        finalTable['Статус налога'][lastIndex] = 'taxable'
        finalTable['В наличии?'][lastIndex] = 1
        finalTable['Возможен ли предзаказ?'][lastIndex] = 0
        finalTable['Продано индивидуально?'][lastIndex] = 0
        finalTable['Разрешить отзывы от клиентов?'][lastIndex] = 1
        finalTable['Категории'][lastIndex] = categories
        finalTable['Изображения'][lastIndex] = picture+', '+picture
        finalTable['Позиция'][lastIndex] = 0
        finalTable['Имя атрибута 1'][lastIndex] = 'Магазин'
        finalTable['Значение(-я) атрибута(-ов) 1'][lastIndex] = 'ул. Плотникова 4, ул. Свободы 67, ул. Страж Революции 4, ул. Телеграфная 53'
        finalTable['Видимость атрибута 1'][lastIndex] = 1
        finalTable['Глобальный атрибут 1'][lastIndex] = 1

        #first child product
        finalTable['ID'][lastIndex+1] = lastIndex+1
        finalTable['Тип'][lastIndex+1] = 'variation'
        # finalTable['Артикул'][lastIndex+1] = SKU
        finalTable['Имя'][lastIndex+1] = name + ' - ' + shop
        finalTable['Опубликован'][lastIndex+1] = 1
        finalTable['рекомендуемый?'][lastIndex+1] = 0
        finalTable['Видимость в каталоге'][lastIndex+1] = 'visible'
        finalTable['Статус налога'][lastIndex+1] = 'taxable'
        finalTable['Налоговый класс'][lastIndex+1] = 'parent'
        finalTable['В наличии?'][lastIndex+1] = 1
        finalTable['Возможен ли предзаказ?'][lastIndex+1] = 0
        finalTable['Продано индивидуально?'][lastIndex+1] = 0
        finalTable['Разрешить отзывы от клиентов?'][lastIndex+1] = 0
        finalTable['Класс доставки'][lastIndex+1] = shop
        if sale_price!=-1:
            finalTable['Цена распродажи'][lastIndex+1] = sale_price
        finalTable['Базовая цена'][lastIndex+1] = price
        # finalTable['Категории'][lastIndex+1] = categories
        # finalTable['Изображения'][lastIndex+1] = picture + ', ' + picture
        finalTable['Класс доставки'][lastIndex+1] = shop
        finalTable['Родительский'][lastIndex+1] = finalTable['Артикул'][lastIndex]
        finalTable['Позиция'][lastIndex+1] = i + 1
        finalTable['Имя атрибута 1'][lastIndex+1] = 'Магазин'
        finalTable['Значение(-я) атрибута(-ов) 1'][lastIndex+1] = shop
        # finalTable['Видимость атрибута 1'][similarity+1+i] = 1
        finalTable['Глобальный атрибут 1'][lastIndex+1] = 1

        return


if __name__ == '__main__':
    myTable = tableMaker()
    id=0
    atribut= ('ул. Плотникова 4', 'ул. Свободы 67', 'ул. Страж Революции 4', 'ул. Телеграфная 53')

    for i in range(len(args.tables)):
        shop=atribut[i]
        df=pd.DataFrame(pd.read_excel(args.tables[i]))
        xlsBook=xlrd.open_workbook(args.tables[i], formatting_info=True)
        sheet=xlsBook[0]
        #print(df)
        #reading rows

        row=11
        for row in range(len(df)):
            #check group
            color=groupChecker(xlsBook,row)
            if color==0:
                group=sheet[row][1].value
                subgroup=''
            if color==1:
                subgroup=sheet[row][1].value
            if color==2:
                price=sheet[row][2]
                tableFiller(myTable, id, sheet[row][1], price, -1, group, subgroup, shop)

    time=datetime.now()
    fileName = 'prices_woo_'+time.strftime('%Y-%m-%d_%H-%M-%S')+'.csv'
    myTable.to_csv(fileName, index=False)



