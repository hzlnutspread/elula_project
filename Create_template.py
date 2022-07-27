import openpyxl
from openpyxl import *
import pandas as pd

paths = ['2022_data', '2021_data', '2020_data', '2019_data']
for path in paths:
    workbook = openpyxl.load_workbook(f'C:/Users/ken/Desktop/Elula_Project/{path}.xlsx', read_only=True, data_only=True)
    print("workbook has been opened")
    # contains all sheetnames)
    sheetnames = workbook.sheetnames
    print(sheetnames)
    workbook.close()

    final_path = f'C:/Users/ken/Desktop/Elula_Project/{path}_frame.csv'
    for dates in sheetnames:
        df = pd.read_csv('C:/Users/ken/Desktop/Elula_Project/Test_data.csv')
        df['Date'] = df['Date'].replace(['15-Oct-21'], f'{dates}')
        df.to_csv(final_path, mode='a', index=False, header=False)
        print(df)

