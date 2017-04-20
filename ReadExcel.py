import pandas as pd
file ='testBook.xlsx'
x1=pd.ExcelFile(file)
print (x1.sheet_names)
#create dataframe from sheet
#the same is
#pd.read_excel(x1,'prvi')
df=x1.parse('prvi')

#gittest
print (df)