# -*- coding: utf-8 -*-
"""
Created on Thu May 11 20:15:57 2017

@author: alex
"""

import pandas as pd

from panda import get_dep_data

book = pd.read_excel('MC.xls',
                     skiprows=[0]) #open local file as dataframe
'''
All MC book data preparation starts HERE:
'''
book = book[book.Сумма != 0] #clear zeros
del book['Unnamed: 5'] #delete no need column

book = book.rename(columns={
'Код и наименование артикула': 'code',
'Ед.изм.': 'name',
'Кол-во': 'quant',
'Сумма': 'summMC',
'Unnamed: 4': 'summPerItem'}) # rename columns

indexes = [] #start indexes of each outlet departmen MC
for i in range(len(book.name)):
    if book.iloc[i][0] == 'Наименование склада: ':
        indexes.append(i)

book_list = [] #list of dataFrames of MC departments
for i in range(len(indexes) - 1):
     books = book.iloc[indexes[i]:indexes[i + 1]]
     book_list.append(books)

if len(book_list) == 20: #check outlets quantity
    print('*' * 20)
    print('Аутлетов 20')
    print('*' * 20)
else:
    print('*' * 20)
    print("Аутлетов %s" % len(book_list))
    print('*' * 20)
'''
End of preparation
'''  
def concaten(depCode): # connect MC outlets in one department
    lists = []    
    for i in range(len(book_list)):
        if book_list[i].values[0][1][-7:] == depCode:
            lists.append(book_list[i])
    return lists




foodList= concaten('1104299') #get dapartment dataFrames list
roomsList = concaten(' 404299')

rooms = pd.concat(concaten(' 404299'))#concatenate dataFrames
food = pd.concat(concaten('1104299'))
eng = pd.concat(concaten(' 764299'))

print(type(foodList[1]))

def clear_all_data_MC(dataFrame): #clear data Function
    '''
    Receive dataFrame as arg.
    on Return we have accounts with sum column only.
    '''
    dataFrame = dataFrame[dataFrame.code == 'Группа ТМЦ: '] #get only accounts data
    del dataFrame['code'] # delete code column
    del dataFrame['quant'] #delete items quantity column
    del dataFrame['summPerItem'] #delete unused columns
    dataFrame['account'] = dataFrame['name'].apply(lambda x: int(x[-4:])*10000) #convert acc column data
    dataFrame['summMC'] = dataFrame['summMC'].apply(lambda x: round(x,2))# convert sum column data
    dataFrame = dataFrame.groupby('account').sum() #sum same accounts
    return dataFrame    

def dep_acc_sums_MC(depData): # get department accounts sum 
    '''
    account departments format:
                dep Name
        account | sum
    '''
    depName = depData.name.values[0][:6]   #get department name 
    depData = clear_all_data_MC(depData) #clear department data
    depData[depName] = depData['summMC'] #change column name -
    del depData['summMC']                # to department name
    return depData

def dep_frames_list(frames_list):
    food_dep_accs = []
    for i in range(len(frames_list)):
        food_dep_accs.append(dep_acc_sums_MC(frames_list[i]))
    return food_dep_accs


#acc_names = get_dep_data(110)[['account','account_name']]
#
#print(get_dep_data(110))
    
SAPfood = get_dep_data(110).groupby(['account']).sum()#.sum() #sum food accounts from SAP
SAProoms = get_dep_data(40).groupby(['account']).sum()
SAPeng = get_dep_data(76).groupby(['account']).sum()
SAPspa = get_dep_data(64).groupby(['account']).sum()
SAPadmin = get_dep_data(72).groupby(['account']).sum()
SAPit = get_dep_data(75).groupby(['account']).sum()
SAPsales = get_dep_data(74).groupby(['account']).sum()
SAPowners = get_dep_data(95).groupby(['account']).sum()


print(SAPfood)
#print(SAPfood.groupby(['account', 'account'], as_index=False).mean())



def compare_dataFrames(SAPdata, MCdata, framesList):
    result = pd.concat([SAPdata, clear_all_data_MC(MCdata)], axis=1) #concatenate SAP+MC data
    MCdeps = pd.concat(dep_frames_list(framesList), axis = 1)
    result = pd.concat([result, MCdeps], axis = 1) #concatenate deps MC frames to SAP dataframe
    result['diff'] = result['summSAP'] - result['summMC'] #make difference column DataMc-DataSAp
    result = result.fillna(value=0)
    result['diff'] = result['diff'].apply(lambda x: round(x)) # round difference to 0 decimals
    result = result[result['summSAP'] != 0] # clear 0 Accounts SAP
    return result

compare_dataFrames(SAPfood, food, foodList)

#result = compare_dataFrames(SAPfood, food, foodList)

#print(pd.concat([result,acc_names], axis=1))

#print(pd.concat([result,acc_names], axis=1))
#print(clear_all_data_MC(food))
#print(SAPfood)

#del food['code'] #deleete first MC column for better excel view

def  wr_to_excel(compare_dataFrames, lists):
    sheet_name = str(int(compare_dataFrames().department.values[0])) #get sheet name

    writer = pd.ExcelWriter('report.xlsx', engine='xlsxwriter') #create writer object
    compare_dataFrames.to_excel(writer, sheet_name=sheet_name) #write Diff data to report
    lists.to_excel(writer, sheet_name=sheet_name, 
              startrow=len(compare_dataFrames)+2, index=False) #write MC data to report

    workbook = writer.book #get book for apply styles
    worksheet = writer.sheets[sheet_name] #get sheet for apply styles

    worksheet.set_column('A:A', 50) #make frist colun 50pix lenght
    worksheet.set_column('C:D', 15) # make rest columns 15 pixels lenght

    writer.save()

wr_to_excel(compare_dataFrames, rooms)







