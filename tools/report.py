import xlwt
import numpy as np
import pandas as pd

path = 'Target.txt'
text = open(path,'r')
clean = text.read().split('nike_ship1')
book = xlwt.Workbook()
sheet1 = book.add_sheet('rejected')
k = 0
for i in range(1, len(clean)):
    item = clean[i].split('\n')
    item_num = (len(clean[i].split('\n')))
    col_name = list(item[2])[29] + list(item[2])[30]

    for q in range(4,item_num):
        row = item[q]
        each_item = row.split('\t')
        clear_each_item = each_item[0].strip( )
        final = clear_each_item.split('\n')
        final_item = final[0].split(' ')
        if final_item[0] != '' and final_item[0] != 'Total':
            sheet1.write(k,0,final_item[0])
            sheet1.write(k,1,int(final_item[1]))
            sheet1.write(k,2,int(final_item[2]))
            sheet1.write(k,3,int(col_name))
            k =k + 1

book.save('test_result.xls')
table = pd.read_excel('test_result.xls',header=None)
p_table_discharged = pd.pivot_table(table,values= 1, index= 0, columns=3 ,aggfunc=np.sum)
p_table_rejected = pd.pivot_table(table,values= 2, index= 0, columns=3 ,aggfunc=np.sum)
p_table_discharged.to_csv('final_discharged.csv')
p_table_rejected.to_csv('final_rejected.csv')