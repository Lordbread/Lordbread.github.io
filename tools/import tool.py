import xlwt

path = 'Target.txt'
text = open(path,'r')
clean = text.read().split('nike_ship1')
book = xlwt.Workbook()

for i in range(1, len(clean)):
    item = clean[i].split('\n')
    item_num = (len(clean[i].split('\n')))
    sheet_name = list(item[2])[29] + list(item[2])[30] + '-'+ list(item[2])[53] +list(item[2])[54]
    sheet1 = book.add_sheet(sheet_name)
    k = 0
    for j in range(3):
        sheet1.write(k,0,item[j])
        k += 1
    for q in range(3,item_num):
        row = item[q]
        each_item = row.split('\t')
        clear_each_item = each_item[0].strip( )
        final = clear_each_item.split('\n')
        final_item = final[0].split(' ')
        if final_item[0] != "":
            if final_item[2] == '(Sum)':
                sheet1.write(q, 3, int(final_item[3]))
                sheet1.write(q, 1, int(final_item[1]))
                sheet1.write(q, 0, final_item[0])
                sheet1.write(q, 2, final_item[2])
            elif final_item[0] == 'Chute':
                sheet1.write(q, 0, final_item[0])
                sheet1.write(q, 1, final_item[1])
                sheet1.write(q, 2, final_item[2])
            elif final_item[1] == '-':
                sheet1.write(q, 1, final_item[1])
                sheet1.write(q, 2, final_item[2])
            else:
                sheet1.write(q, 0, final_item[0])
                sheet1.write(q, 1, int(final_item[1]))
                sheet1.write(q, 2, int(final_item[2]))
book.save('result.csv')