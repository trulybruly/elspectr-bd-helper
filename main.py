import csv
import argparse
import pandas as pd

#parser = argparse.ArgumentParser(usage="elspectr", description="4 xls files into csv for shop")
#parser.add_argument("-q","first",required=True, help="First file")
#parser.add_argument("-w","--second",required=True, help="Second file")
#parser.add_argument("-e","--third",required=True, help="Third file")
#parser.add_argument("-r","--fourth",required=True, help="Fourth file")
#args = parser.parse_args()

if __name__ == '__main__':
    df=pd.DataFrame(pd.read_excel("/home/trulybruly/Загрузки/prices/price_varya.xls"))
    df
