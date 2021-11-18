import csv
import argparse
import pandas as pd
import xlrd

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
                            'Родительский':[],'Сгруппированные товары':[],'Апсейл':[],'Кросселы':[],'Внешний URL':[],'Текст кнопки':[],'Позиция':[],'Имя атрибута 1':[],'Значение(-я) аттрибута(-ов) 1':[],
                            'Видимость атрибута 1':[],'Глобальный атрибут 1':[]})
    return myTable


def tableFiller(finalTable,id,name,price,sale_price,group,subgroup,):
    attributeName='Магазин'



if __name__ == '__main__':
    myTable = tableMaker()
    id=1
    attribute= ('ул. Плотникова 4','ул. Свободы 67','ул. Страж Революции 4','ул. Телеграфная 53')

    for i in range(len(args.tables)):
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
                tableFiller()



