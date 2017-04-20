#http://stackoverflow.com/questions/23568409/xlrd-python-reading-excel-file-into-dict-with-for-loops
#http://stackoverflow.com/questions/2725852/writing-to-existing-workbook-using-xlwt
#http://stackoverflow.com/questions/3723793/preserving-styles-using-pythons-xlrd-xlwt-and-xlutils-copy

#import xlwt
import xlrd
import glob
from xlutils.copy import copy

first_path='test1.xls'
sec_path='test2.xls'
for name in glob.glob("test*.xls"):
    print(name)

rb = xlrd.open_workbook(first_path,formatting_info=True)
r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
wb = copy(rb) # a writable copy (I can't read values out of this, only write to it)
w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy

colsToCopy=r_sheet.ncols
rowToStart = r_sheet.nrows
print("rows 1:", rowToStart, "  Colls:", colsToCopy)


r2 = xlrd.open_workbook(sec_path,formatting_info=True)
r_sheet2 = r2.sheet_by_index(0)
for row_idx in range(1,r_sheet2.nrows):
    for col_idx in range(0,colsToCopy):
        cell = r_sheet2.cell(row_idx,col_idx)
        w_sheet.write(rowToStart, col_idx, cell.value)
    rowToStart +=1

wb.save('joined.xls')



'''
xlsfiles = [r'SlavisaIzvestaj2.xls', r'SlavisaIzvestaj3.xls']
# read header values into the list
sheetStart = xlrd.open_workbook(r'SlavisaIzvestaj1.xls').sheets()[0]
outsheet = wkbk.add_sheet(sheetStart)
print("rows Start:", sheetStart.nrows, "  Colls:", sheetStart.ncols)
keys = [sheetStart.cell(0, col_index).value for col_index in xrange(sheetStart.ncols)]

outrow_idx = 0
for f in xlsfiles:
    insheet = xlrd.open_workbook(f).sheets()[0]
    print("rows:", insheet.nrows, "  Colls:", insheet.ncols)
    for row_idx in xrange(insheet.nrows):
        for col_idx in xrange(insheet.ncols):
            outsheet.write(outrow_idx, col_idx,
                           insheet.cell_value(row_idx, col_idx))
        outrow_idx += 1
wkbk.save(r'combined.xls')
'''