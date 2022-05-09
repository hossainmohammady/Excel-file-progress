import xlrd
import xlwt
import os
inwb = xlwt.Workbook()
inws = inwb.add_sheet('Auto', cell_overwrite_ok=True)
for i, file in enumerate(os.listdir()):
    if file.endswith(".xls") and file != ".Auto.xls":
        wb = xlrd.open_workbook(file)
        sh = wb.sheet_by_index(0)
        inws.write(i, 1, sh.cell_value(1, 0))  # جمع کمیسیون
        inws.write(i, 0, sh.cell_value(5, 0))  # تاریخ
        inws.write(i, 2, sh.cell_value(1, 1))  # نسیه برگشتی
        inws.write(i, 3, sh.cell_value(1, 2))  # جمع کارتخوان
        inws.write(i, 4, sh.cell_value(1, 3))  # هزینه
        inws.write(i, 5, sh.cell_value(1, 4))  # نسیه
        inws.write(i, 6, sh.cell_value(6, 5))  # نقدی
        print(f"End of progress file: [ {file} ]")

print(f"End of progress all file in folder  Please check file .Auto.xls")
inwb.save('.Auto.xls')
#print("Do you want to open .auto.xls file? (y/n)")
questin = input()
# if questin=="y":
#  os.open('.Auto.xls', 1)
