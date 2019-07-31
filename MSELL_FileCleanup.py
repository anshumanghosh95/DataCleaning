# -*- coding: utf-8 -*-
"""
Created on Thu Jul 25 20:03:52 2019

@author: AN389897
"""
import pandas as pd
import numpy as np
import os

try:
    os.chdir(os.path.curdir)
    data = pd.read_excel('Huawei Y7 Pro Nova 3i P30 P30 Pro Oppo F11 Pro Marvel Asus ROG S10e B1F1....xlsx', sheetname = 'Config - K2 (2)')
    data_header = pd.read_excel('Huawei Y7 Pro Nova 3i P30 P30 Pro Oppo F11 Pro Marvel Asus ROG S10e B1F1....xlsx', sheetname = 'Config - K2 (2)', header = None)
    list_columns = list(data_header.iloc[0,:].values)
    list_columns = list(map(lambda x : x.replace('\n',' '),list_columns))
    data.columns = data.columns.str.replace('\n',' ')
    data = data[data['UOM/ Sales Unit'] != 'EA']
    for i, key in data.iterrows():
        if np.isnan([key['Article']]) == [True]:
            data = data.drop(i)
    data['Article Description'] = data.apply(lambda X: X[1].replace(',','.'), axis = 1)
    data['Trade Up Amount (RM)'] = data.apply(lambda X: X[-1].replace(',','.'), axis = 1)
    
    dict_string = {'Article': 0, 'Account  Category': 5, 'Market Code': 7, 'Contract Tenure': 9, 'Kenan  Package ID': 10, 'Component ID': 11,'Kenan Contract Component': 13}
    for i,k in dict_string.items():
        data[i] = data.apply(lambda X: str(X[k]).replace('.0',''), axis = 1)
    
    for i, key in data.iterrows():
        if [key['Component ID']] == ['nan']:
            data['Component ID'] = data.apply(lambda X: str(X[11]).replace('nan',''), axis = 1)
    data['Component ID'] = data['Component ID'].astype(object)
    
    #Changing date format
    data['BBTU End Date [mm/dd/yyyy]'] = data['BBTU End Date [mm/dd/yyyy]'].astype(str).replace('\.0', '', regex=True)
    data['BBTU End Date [mm/dd/yyyy]'] = data['BBTU End Date [mm/dd/yyyy]'].astype(str).replace('9999-12-31 00:00:00','12/31/9999')
    data['End date [mm/dd/yyyy]'] = data['End date [mm/dd/yyyy]'].astype(str).replace('\.0', '', regex=True)
    data['End date [mm/dd/yyyy]'] = data['End date [mm/dd/yyyy]'].astype(str).replace('9999-12-31 00:00:00','12/31/9999')
    data['Start date [mm/dd/yyyy]'] = pd.to_datetime(data['Start date [mm/dd/yyyy]']).dt.strftime('%m/%d/%Y')
    data['BBTU Start Date [mm/dd/yyyy]'] = pd.to_datetime(data['BBTU Start Date [mm/dd/yyyy]']).dt.strftime('%m/%d/%Y')
    
    #Changing RM format
    data['Plan Price (RM) [00.00]'] = data.apply(lambda X: format(X[3],'.2f'), axis = 1)
    data['Mass A Foreigner/ Mass C Device Avance Payment (RM)'] = data.apply(lambda X: format(X[-8],'.2f'), axis = 1)
    data['Rate plan advance payment  (RM - No Longer Applicable for Device Contract Bundles) [00.00]'] = data.apply(lambda X: format(X[-7],'.2f'), axis = 1)
    data['Deposit  (Malaysians)  (RM) [00.00]']  = data.apply(lambda X: format(X[-6],'.2f'), axis = 1)
    data['Deposit  (Others)  (RM) [00.00]'] = data.apply(lambda X: format(X[-5],'.2f'), axis = 1)
    
    data.loc[0] = list_columns
    data = data.sort_index()
    data.to_csv('MSELL.csv',sep = ',', header = None, index = False)
except Exception as e:
    print(e)