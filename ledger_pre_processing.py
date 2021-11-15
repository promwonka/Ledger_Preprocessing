import os
import pandas as pd
import numpy as np
import glob

import sys
 

cwd = os.getcwd()

def getIndexes(dfObj, value):
    listOfPos = list()
    result = dfObj.isin([value])
    seriesObj = result.any()
    columnNames = list(seriesObj[seriesObj == True].index)
    for col in columnNames:
        rows = list(result[col][result[col] == True].index)
        for row in rows:
            listOfPos.append((row, col))
    return listOfPos





def ledger_pre():

    os.chdir(cwd)
    extension = 'xlsx'
    all_filenames = [i for i in glob.glob('*.{}'.format(extension))]
    list_file=[]

    print(all_filenames)
    for files in all_filenames:
            dataset= pd.read_excel(files)#,encoding = "unicode_escape")
            dataset.columns =['Date','particulars','sales', 'vch type','vch_number','debit','credit']
            listOfPositions = getIndexes(dataset, 'Ledger:')
            k=listOfPositions[0][0] 
            dataset = dataset.iloc[k:]
            dataset['customer_name']=dataset['particulars']
            dataset=dataset.replace(to_replace ="To", value =np.nan) 
            dataset=dataset.replace(to_replace ="By", value =np.nan) 
            dataset=dataset.replace(to_replace ="Particulars", value =np.nan) 
            pd.set_option('display.max_rows',None)
            pd.set_option('display.max_columns',None)
            cols = ['customer_name']
            dataset.loc[:,cols] = dataset.loc[:,cols].ffill()
            dataset = dataset[dataset.Date != 'Ledger:']
            dataset = dataset[dataset.Date !='Date']
            dataset=dataset.drop(['particulars'], axis = 1)
            dataset['credit'] = dataset['credit'].fillna(0)
            dataset['debit'] = dataset['debit'].fillna(0)
            dataset=dataset.drop(['vch type'], axis = 1) 
            dataset=dataset.drop(['vch_number'], axis = 1)
            dataset['Date'] = pd.to_datetime(dataset['Date'], errors='coerce')
            dataset=dataset.dropna(subset=['sales'])
            cols = ['Date']
            dataset.loc[:,cols] = dataset.loc[:,cols].ffill()
            dataset.reset_index(inplace = True, drop = True)
            list_file.append(dataset)
    result = pd.concat(list_file)    
    result.sort_values(by=['customer_name', 'Date'])
    for value in result['sales']:
        if value == 'Opening Balance':
            result.drop_duplicates(subset=['customer_name', 'sales'])
    result = result[result.sales != 'Closing Balance']
    result = result.rename(columns={"Date":"date"})
    result.to_csv("ledger_all.csv", index = False)
    
ledger_pre()