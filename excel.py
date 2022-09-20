import os
import glob
import csv
from xlsxwriter.workbook import Workbook

#Iterates over every CSV file in the folder and saves an XLSX version.
for csvfile in glob.glob(os.path.join(r'\Users\asiegel\Downloads\Checkin Program\csv', '*.csv')):
    workbook = Workbook(csvfile[:-4] + '.xlsx')
    print(csvfile[:-4])
    worksheet = workbook.add_worksheet()
    
    with open(csvfile, 'rt', encoding='utf8') as f:
        
        reader = csv.reader(f)

        #Formatting
        format =  workbook.add_format()
        format.set_border()
           
        #Write Data   
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
                worksheet.write(r, c, col, format)

                #Set BG Color, C is the Row number
                if c > 3:
                    format_bg =  workbook.add_format()
                    format_bg.set_border()
                    format_bg.set_bg_color('#F4B084')
                    worksheet.write(r, c, col, format_bg)

        #Creates table for each Excel File, Makes Length specific to each
        worksheet.add_table(0,0,r,5, {'style' :'Table Style Medium 2',
        'columns' :[
                {'header': 'PO'},
                {'header': 'PO Type'},
                {'header': 'Ship Date'},
                {'header': 'Newness flag'},
                {'header': 'Vendor Confirmed Ship Date'},
                {'header': 'Comments'},
                ]})      
                  
        #AutoFit Work Around, set each column length        
        worksheet.set_column(0, 0, 10)
        worksheet.set_column(1, 1, 15)
        worksheet.set_column(2, 2, 11)
        worksheet.set_column(3, 3, 12)
        worksheet.set_column(4, 4, 26)
        worksheet.set_column(5, 5, 10)            
    workbook.close()