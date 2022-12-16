# Import libraries required
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import xlwings
import openpyxl
import time

# Record processing time
start_time = time.time()


# Input the file name
print('Input the file name')
File_Date=str(input())
File_name=File_Date+'.xls'
print(File_name)

# Extract data from file
Volume_Origianl = pd.read_csv(File_name)
Volume_Origianl=Volume_Origianl.reset_index()
Volume_Origianl = Volume_Origianl.drop(Volume_Origianl.columns[2],axis = 1)

# Create and format empty dataframe, For 'DateTime','ID','Length','Width','Height','Equipment_ID'
df = pd.DataFrame(columns=['DateTime','ID','Length','Width','Height','Equipment_ID','FSALDU','Service','Barcode_Type','FSA','Destination','FeedDate','FeedTime','Maxlength','Hour','Volume','Size','Shift','Fine_Dest'])
for i in range(6):
 df[df.columns[i]]=Volume_Origianl.iloc[:,i]

# For 'Barcode type' and FSA/LDU if applicable
for x in df.itertuples():
 if len(x.ID)>16 and len(x.ID)<22:
    df.at[x.Index,'Barcode_Type']== 'SYN'
 if len(x.ID)==16:
    df.at[x.Index,'Barcode_Type']='292'
 if len(x.ID)>22:
    df.at[x.Index,'Barcode_Type']='NGB'
    
for x in df.itertuples():
 if x.Barcode_Type=='NGB':
    df.at[x.Index,'FSALDU']=x.ID[4]+x.ID[7]+x.ID[5]+x.ID[8]+x.ID[6]+x.ID[9]

df['FSA']=df['FSALDU'].str.slice(stop=3)

print('Report In Progress')

# Write FSA to destination excel file for dest lookup
with pd.ExcelWriter('Destination.xlsx', mode='a',if_sheet_exists='replace') as writer:  
 df.to_excel(writer, sheet_name='FSA',columns=['FSA'])
excel_app = xlwings.App(visible=False)
excel_book = excel_app.books.open('Destination.xlsx')
excel_book.save()
excel_book.close()
excel_app.quit()
df2 = pd.read_excel('Destination.xlsx', sheet_name=2, nrows=len(Volume_Origianl))
df['Destination']=df2['Dest']

# Finer destination
df3 = pd.read_csv('Fine_Dest.csv',nrows=df.shape[0])
# Finer destination-Service
df4=df3.iloc[0: 14 ,0: 5]
df4=df4.set_index('Dest')
# Finer destination-FSALDU 
df5=df3.iloc[0: 18 ,7: 9]
df5 = df5.set_index('Postal_Code')

# For 'FeedDate','FeedTime','Hour','Volume'
df['FeedDate']=df['DateTime'].str.slice(stop=10)
df['FeedTime']=df['DateTime'].str.slice(start=-8)
df['Hour']=df['FeedTime'].str.slice(start=0,stop=2)
df['Volume']=round(df['Length']*df['Width']*df['Height']/16387.064)
df=df.astype({'Hour':int,'Length':int,'Width':int,'Height':int,'Volume':int,'Size':str})

# Iterate each row to update value for 'Service','Shift','Maxlength','Size','Destination'
for x in df.itertuples():
# For service type
 if x.Barcode_Type=='292':
    df.at[x.Index,'Service']=x.ID[3]
 if x.Barcode_Type=='NGB':
    df.at[x.Index,'Service']=x.ID[3]
# For Shift
 if x.Hour>6 and x.Hour<15:
    df.at[x.Index,'Shift']=2
 elif x.Hour>=15 and x.Hour<23:
    df.at[x.Index,'Shift']=3
 else:
    df.at[x.Index,'Shift']=1
# For Maxlength
 if x.Length>x.Width and x.Length>x.Height:
    df.at[x.Index,'Maxlength']=x.Length/25.4
 elif x.Width>x.Height:
    df.at[x.Index,'Maxlength']=x.Width/25.4
 else:
    df.at[x.Index,'Maxlength']=x.Height/25.4

# For destination
 if x.Destination in df4.index:
   if x.Service=='P':
        df.at[x.Index,'Fine_Dest']=df4.loc[x.Destination,str(x.Service)] 
   elif (x.Service=='1') | (x.Service=='2') |(x.Service=='3'):
        df.at[x.Index,'Fine_Dest']=df4.loc[x.Destination,str(x.Service)]
 elif x.FSALDU in df5.index:
      df.at[x.Index,'Fine_Dest']=df5.loc[x.FSALDU,['Dest2']]
 else:  
      df.at[x.Index,'Fine_Dest']=x.Destination

# For size
for x in df.itertuples():
 if x.Maxlength==0:
  df.at[x.Index,'Size']='No Data'
 if x.Maxlength>0 and x.Maxlength<=12:
  df.at[x.Index,'Size']='0-12"'
 if x.Maxlength>12 and x.Maxlength<=24:
  df.at[x.Index,'Size']='12-24"'
 if x.Maxlength>24 and x.Maxlength<=36:
  df.at[x.Index,'Size']='24-36"'
 if x.Maxlength>36:
  df.at[x.Index,'Size']='>36"'

# For 'Equipment_ID'
df['Equipment_ID'] = df['Equipment_ID'].replace([1260],'CUBI1')
df['Equipment_ID'] = df['Equipment_ID'].replace([1261],'CUBI3')
df['Equipment_ID'] = df['Equipment_ID'].replace([1262],'CUBI2')
df['Equipment_ID'] = df['Equipment_ID'].replace([1263],'CUBI4')
df['Equipment_ID'] = df['Equipment_ID'].replace([1264],'CUBI5')

#Generate report
df.to_csv('Daily_report_Python.csv') 
print('Report Complete')
print("--- %s seconds ---" % (time.time() - start_time))