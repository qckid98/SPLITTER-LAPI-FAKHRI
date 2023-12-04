# splitting machine
import pandas as pd
import xlwings as xw

# create workabook
wb =  xw.Book(r"C:\Users\Asrock-PC\Downloads\Split.xlsx")
ws = wb.sheets['Split']

# lumping data
lumping = pd.read_excel(r"C:\Users\Asrock-PC\Downloads\Split.xlsx", sheet_name = 'Res (K x H)')
# perforation data
perfo = pd.read_excel(r"C:\Users\Asrock-PC\Downloads\Split.xlsx", sheet_name = 'Production & Open Completion')
# well data
well = pd.read_excel(r"C:\Users\Asrock-PC\Downloads\Split.xlsx", sheet_name = 'Well')

row_count = 2
i = 0

for index in well.index:
    
    while i < len(perfo):
        
        if str(perfo['UWI'][i]) == str(well['Well'][index]):
            
            # define total_kh initial
            total_kh = 0
            # define formation count intitial
            perfo_count = 16

            for kh in lumping.index:
                if perfo.iloc[row_count-2, perfo_count] == 'P':
                    total_kh += lumping[str(well['Well'][index])][kh]
                perfo_count += 1

            perfo_count = 16
            num = 14

            # num pada air dan gas harus disesuaikan dengan jumlah sandnya
            for form in lumping.index:
                if total_kh > 0:
                    if perfo.iloc[row_count-2, perfo_count] == 'P':
                        ws.range(row_count+1, num).value = float(perfo['CUM OIL'][i])*float(lumping[str(well['Well'][index])][form])/total_kh
                        ws.range(row_count+1, num+28).value = float(perfo['CUM WATER'][i])*float(lumping[str(well['Well'][index])][form])/total_kh
                        ws.range(row_count+1, num+56).value = float(perfo['CUM GAS'][i])*float(lumping[str(well['Well'][index])][form])/total_kh
                    else:
                        ws.range(row_count+1, num).value = 0
                        ws.range(row_count+1, num+28).value = 0
                        ws.range(row_count+1, num+56).value = 0
                perfo_count += 1
                num += 1  

            row_count += 1
            i += 1

        else:
            break