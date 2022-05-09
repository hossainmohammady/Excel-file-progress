import xlrd
import xlwt
import os
inwb = xlwt.Workbook()  # open the workbook for exporting
# open the sheet exporting
inws = inwb.add_sheet('Auto', cell_overwrite_ok=True)
for i, file in enumerate(os.listdir()):  # loop through the files in the directory
    # if the file is an excel file and is not the auto file
    if file.endswith(".xls") and file != ".Auto.xls":
        wb = xlrd.open_workbook(file)  # open the excel file for im
        sh = wb.sheet_by_index(0)  # select the first sheet
        # select cells in the first sheet for exporting
        inws.write(i, 1, sh.cell_value(1, 0))  # جمع کمیسیون
        inws.write(i, 0, sh.cell_value(5, 0))  # تاریخ
        inws.write(i, 2, sh.cell_value(1, 1))  # نسیه برگشتی
        inws.write(i, 3, sh.cell_value(1, 2))  # جمع کارتخوان
        inws.write(i, 4, sh.cell_value(1, 3))  # هزینه
        inws.write(i, 5, sh.cell_value(1, 4))  # نسیه
        inws.write(i, 6, sh.cell_value(6, 5))  # نقدی
        print(f"End of progress file: [ {file} ]")  # print the file name

inwb.save('.Auto.xls')
print(f"End of progress all file in folder  Please check file .Auto.xls")
#print("Do you want to open .auto.xls file? (y/n)")
questin = input()  # for dont close the program
# if questin=="y":
#  os.open('.Auto.xls', 1)
