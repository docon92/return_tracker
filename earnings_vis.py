# ------ earnings_vis.py ------------
#Pull TFSA, RRSP Data from an Excel spreadsheet and visualize contributions, earnings, etc.

import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import openpyxl


# TODO: load the file from script argument
# Load the excel file in this dir, extract data on both RRSP,TFSA Sheets
df = openpyxl.load_workbook(r'template_accounts.xlsx',data_only=True)
TFSA = df['TFSA']
RRSP = df['RRSP']

#print(TFSA.max_row)
#TODO: Detect actual end of file. TFSA.max_row will give you
#      something larger than reality if there are blank cells
#
# Note that the first row is given to column titles, Data starts at row 2

ROW_START = 2
MAX_ROW = 33

#  Initialize Datasets for plotting

TFSA_val = []
RRSP_val = []
TFSA_cont = []
RRSP_cont= []
Total_val = []
Total_cont = []
Total_earn = []

for row in range(ROW_START, MAX_ROW):
    B_row = 'B' + str(row)
    C_row = 'C' + str(row)
    #print('index is: %s ' % index )
    Data1 = TFSA[B_row].value
    Data2 = RRSP[B_row].value
    Data3 = TFSA[C_row].value
    Data4 = RRSP[C_row].value
    TFSA_val.append(Data1)
    RRSP_val.append(Data2)
    TFSA_cont.append(Data3)
    RRSP_cont.append(Data4)
    Total_val.append(Data1+Data2)
    Total_cont.append(Data3+Data4)
    Total_earn.append((Data1+Data2)-(Data3+Data4))
    print('MultiSheet Data: %s , %s , %s , %s' % (Data1, Data2,Data3,Data4))

# Make x axis length in years
x=[]
for i in range(MAX_ROW-ROW_START):
    x.append(i/12)

# Stack Plot Total contributions and total earnings
y = [Total_cont,Total_earn]

#plt.legend(loc='upper left')
sns.set_style(u'whitegrid')
# Change the color and its transparency
#plt.fill_between( x, y1 , color="skyblue", alpha=0.6)
#plt.fill_between( x, y2, color="orange", alpha=0.2)

pal = sns.color_palette("Set2")
plt.stackplot(x,y,colors=pal, alpha=0.8 )

#plt.tick_params(labelbottom='off')
#plt.tick_params(labelleft='off')
plt.show()




