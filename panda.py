# -*- coding: utf-8 -*-
"""
Created on Thu May 11 20:15:57 2017

@author: alex
"""

import pandas as pd

book = pd.read_excel('trial-balance.xls',
                     skiprows=[0], parse_cols=[0,1,3,4,6]) #open local file as dataframe

#print(book)                    
                     
book = book.rename(columns= {
'Unnamed: 0':'reason',
'Unnamed: 1':'department',
'Unnamed: 2':'account',
'Unnamed: 3':'account_name',
0 :'summSAP'})   #rename columns




book = book[book.reason =='AMOUNT'] #clear data
book = book[book.account > 70000000]
book = book[book.summSAP != 0.00]
del book['reason']

#deps = (40,110,76,72,75,95,64)
#print(book)

def get_dep_data(depNumber):
    '''
    Get department data by department number
    
    We have DataFrame out
    '''
    dataFrame = book[book.department == depNumber]
    return dataFrame


SAPfood = get_dep_data(110).groupby(['account']).sum()#.sum() #sum food accounts from SAP
SAProoms = get_dep_data(40).groupby(['account']).sum()
SAPeng = get_dep_data(76).groupby(['account']).sum()
SAPspa = get_dep_data(64).groupby(['account']).sum()
SAPadmin = get_dep_data(72).groupby(['account']).sum()
SAPit = get_dep_data(75).groupby(['account']).sum()
SAPsales = get_dep_data(74).groupby(['account']).sum()
SAPowners = get_dep_data(95).groupby(['account']).sum()

#print(SAPfood)