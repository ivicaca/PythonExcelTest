import pandas as pd
import glob

#filenames to combine
#excel_names=['testBook.xlsx', 'testBook2.xlsx']

allData=pd.DataFrame()
#read excels
for name in glob.glob("test*.xlsx"):
    df=pd.read_excel(name)
    allData=allData.append(df,ignore_index=True)
    print (name)
print (allData.head())

#excels=[pd.ExcelFile(name) for name in excel_names]
