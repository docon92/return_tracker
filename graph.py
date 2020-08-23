import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import openpyxl

#data = pd.ExcelFile (r'/Users/douglasconover/Documents/R-A-Total.xlsx')
df = openpyxl.load_workbook(r'/Users/douglasconover/Documents/R-A-Total.xlsx',data_only=True)
#df.info

#print(df.info)
sheet2 = df['TFSA']
sheet1 = df['RRSP']
print(sheet1.max_row)

y1 = []
y2 = []
y3 = []
y4 = []

data_len = 31

for row in range(2, data_len+2):
    index1 = 'B' + str(row)
    index2 = 'C' + str(row)
    #print('index is: %s ' % index )
    Data1 = sheet1[index1].value
    Data2 = sheet2[index1].value
    Data3 = sheet1[index2].value
    Data4 = sheet2[index2].value
    y1.append(Data1)
    y2.append(Data2)
    temp = (Data1+Data2)-(Data3+Data4)
#    y3.append(temp)
    y4.append(Data1+Data2-Data3-Data4)
    y3.append(Data3+Data4)
    print('MultiSheet Data: %s , %s , %s , %s' % (Data1, Data2,Data3,Data4))
#print(sheet['B18'].value)
# Data
x=range(2,data_len+2)
#y=[ [1,4,6,8,9], [2,2,7,10,12], [2,8,5,10,6] ]

y = [y3,y4]
print(y)
# Plot
#plt.stackplot(x,y, labels=['A','B','C'])
#plt.legend(loc='upper left')
pal = sns.color_palette("Set1")

plt.stackplot(x,y,colors=pal, alpha=0.4 )
plt.show()
