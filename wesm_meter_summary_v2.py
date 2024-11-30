#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#this script creates csv files from WESM monthly meter data files which are excel and are badly summarized
import pandas as pd
from glob import glob


# In[ ]:


print('WESM meter montlhy summary prepared by gpureta github: https://github.com/gpureta/WESM-TRADING-SCRIPTShttps://github.com/gpureta/WESM-TRADING-SCRIPTS')
print('Please copy metering files in the same folder as executable script')
print('Note: SSLA is only "Line Loss" column')


# In[ ]:


#user input the type of interval if hourly or 5min or both
intervalcheck = 0
while (intervalcheck not in ["f","h","b"]):
    print('Enter "f" for five-minute interval summary. Enter "h" for hourly interval summary. Enter "b" for both.')
    interval = input()
    intervalcheck = interval


# In[ ]:


print("running script... press [ctrl + c] to abort")


# In[ ]:


# create list of all xlsx files
meter_files = sorted(glob('*MonthlyMQ.xlsx'))
meter_files


# In[ ]:


#get sheet names
sheetnames = pd.ExcelFile(meter_files[0]).sheet_names


# In[ ]:


#output csv file columns
column_names = ["date","date2","interval"] + sheetnames


# In[ ]:


#script for 5min interval
if interval in ["f"]:
    #creating new dataframes
    RAW = pd.DataFrame(columns = column_names)
    SSLA = pd.DataFrame(columns = column_names)
    ADJUSTED = pd.DataFrame(columns = column_names)
    CAPTIVE = pd.DataFrame(columns = column_names)

    RAW2 = pd.DataFrame(columns = column_names)
    SSLA2 = pd.DataFrame(columns = column_names)
    ADJUSTED2 = pd.DataFrame(columns = column_names)
    CAPTIVE2 = pd.DataFrame(columns = column_names)
    name = '5min'
    x = len(pd.ExcelFile(meter_files[0]).sheet_names)
    for day in meter_files:

        if x!= len(pd.ExcelFile(day).sheet_names):
                print("error: variable meter numbers. check for line switching on",day.split('_')[1])
                print("creating incomplete csv file...")
                break
        RAW["interval"] = pd.date_range(start="00:00:00",end="23:55:00", freq='5min').strftime('%H:%M:%S').tolist()
        RAW["date"] = pd.read_excel(day, sheet_name =sheetnames[0]).iloc[0,1]
        RAW["date2"] = day.split('_')[1]
        ADJUSTED["interval"] = RAW["interval"].copy()
        ADJUSTED["date"] = RAW["date"].copy()
        ADJUSTED["date2"] = RAW["date2"].copy()
        SSLA["interval"] = RAW["interval"].copy()
        SSLA["date"] = RAW["date"].copy()
        SSLA["date2"] = RAW["date2"].copy()
        CAPTIVE["interval"] = RAW["interval"].copy()
        CAPTIVE["date"] = RAW["date"].copy()
        CAPTIVE["date2"] = RAW["date2"].copy()
        for sheet in sheetnames:

            RAW[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 1::6].values.reshape(288)
            ADJUSTED[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 2::6].values.reshape(288)
            SSLA[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 3::6].values.reshape(288)
            CAPTIVE[sheet]=pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 6::6].values.reshape(288)
        RAW2= RAW2.append(RAW)
        SSLA2 = SSLA2.append(SSLA)
        ADJUSTED2 =ADJUSTED2.append(ADJUSTED)
        CAPTIVE2 = CAPTIVE2.append(CAPTIVE)

        print(day,":",len(pd.ExcelFile(day).sheet_names))

        
    RAW2.to_csv('RAW'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    SSLA2.to_csv('SSLA'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    ADJUSTED2.to_csv('ADJUSTED'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    CAPTIVE2.to_csv('CAPTIVE'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    print("five minute csv files created.")


# In[ ]:


#script for hourly interval
if interval in ["h"]:
    #creating new dataframes
    RAW = pd.DataFrame(columns = column_names)
    SSLA = pd.DataFrame(columns = column_names)
    ADJUSTED = pd.DataFrame(columns = column_names)
    CAPTIVE = pd.DataFrame(columns = column_names)

    RAW2 = pd.DataFrame(columns = column_names)
    SSLA2 = pd.DataFrame(columns = column_names)
    ADJUSTED2 = pd.DataFrame(columns = column_names)
    CAPTIVE2 = pd.DataFrame(columns = column_names)
    name = 'hourly'
    x = len(pd.ExcelFile(meter_files[0]).sheet_names)
    for day in meter_files:
        
        if x!= len(pd.ExcelFile(day).sheet_names):
                print("error: variable meter numbers. check for line switching on",day.split('_')[1])
                print("creating incomplete csv file...")
                break

        RAW["interval"] = hour = list(range(1,25))
        RAW["date"] = pd.read_excel(day, sheet_name =sheetnames[0]).iloc[0,1]
        RAW["date2"] = day.split('_')[1]
        ADJUSTED["interval"] = RAW["interval"].copy()
        ADJUSTED["date"] = RAW["date"].copy()
        ADJUSTED["date2"] = RAW["date2"].copy()
        SSLA["interval"] = RAW["interval"].copy()
        SSLA["date"] = RAW["date"].copy()
        SSLA["date2"] = RAW["date2"].copy()
        CAPTIVE["interval"] = RAW["interval"].copy()
        CAPTIVE["date"] = RAW["date"].copy()
        CAPTIVE["date2"] = RAW["date2"].copy()
        for sheet in sheetnames:

            RAW[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 1::6].values.sum(axis= 1)
            ADJUSTED[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 2::6].values.sum(axis= 1)
            SSLA[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 3::6].values.sum(axis= 1)
            CAPTIVE[sheet]=pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 6::6].values.sum(axis= 1)
        RAW2= RAW2.append(RAW)
        SSLA2 = SSLA2.append(SSLA)
        ADJUSTED2 =ADJUSTED2.append(ADJUSTED)
        CAPTIVE2 = CAPTIVE2.append(CAPTIVE)
        
        print(day,":",len(pd.ExcelFile(day).sheet_names))

    RAW2.to_csv('RAW'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    SSLA2.to_csv('SSLA'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    ADJUSTED2.to_csv('ADJUSTED'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    CAPTIVE2.to_csv('CAPTIVE'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    print("hourly csv files created.")


# In[ ]:


#script for 5min interval and hourly
if interval in ["b"]:
    #creating new dataframes for fivemin
    RAW = pd.DataFrame(columns = column_names)
    SSLA = pd.DataFrame(columns = column_names)
    ADJUSTED = pd.DataFrame(columns = column_names)
    CAPTIVE = pd.DataFrame(columns = column_names)

    RAW2 = pd.DataFrame(columns = column_names)
    SSLA2 = pd.DataFrame(columns = column_names)
    ADJUSTED2 = pd.DataFrame(columns = column_names)
    CAPTIVE2 = pd.DataFrame(columns = column_names)
    name = '5min'
    
    #creating new dataframes for hourly
    h_RAW = pd.DataFrame(columns = column_names)
    h_SSLA = pd.DataFrame(columns = column_names)
    h_ADJUSTED = pd.DataFrame(columns = column_names)
    h_CAPTIVE = pd.DataFrame(columns = column_names)

    h_RAW2 = pd.DataFrame(columns = column_names)
    h_SSLA2 = pd.DataFrame(columns = column_names)
    h_ADJUSTED2 = pd.DataFrame(columns = column_names)
    h_CAPTIVE2 = pd.DataFrame(columns = column_names)
    h_name = 'hourly'
    
    x = len(pd.ExcelFile(meter_files[0]).sheet_names)
    for day in meter_files:

        if x!= len(pd.ExcelFile(day).sheet_names):
                print("error: variable meter numbers. check for line switching on",day.split('_')[1])
                print("creating incomplete csv file...")
                break
        #five minute
        RAW["interval"] = pd.date_range(start="00:00:00",end="23:55:00", freq='5min').strftime('%H:%M:%S').tolist()
        RAW["date"] = pd.read_excel(day, sheet_name =sheetnames[0]).iloc[0,1]
        RAW["date2"] = day.split('_')[1]
        ADJUSTED["interval"] = RAW["interval"].copy()
        ADJUSTED["date"] = RAW["date"].copy()
        ADJUSTED["date2"] = RAW["date2"].copy()
        SSLA["interval"] = RAW["interval"].copy()
        SSLA["date"] = RAW["date"].copy()
        SSLA["date2"] = RAW["date2"].copy()
        CAPTIVE["interval"] = RAW["interval"].copy()
        CAPTIVE["date"] = RAW["date"].copy()
        CAPTIVE["date2"] = RAW["date2"].copy()
        for sheet in sheetnames:

            RAW[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 1::6].values.reshape(288)
            ADJUSTED[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 2::6].values.reshape(288)
            SSLA[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 3::6].values.reshape(288)
            CAPTIVE[sheet]=pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 6::6].values.reshape(288)
        RAW2= RAW2.append(RAW)
        SSLA2 = SSLA2.append(SSLA)
        ADJUSTED2 =ADJUSTED2.append(ADJUSTED)
        CAPTIVE2 = CAPTIVE2.append(CAPTIVE)
        
        #hourly 
        h_RAW["interval"] = hour = list(range(1,25))
        h_RAW["date"] = pd.read_excel(day, sheet_name =sheetnames[0]).iloc[0,1]
        h_RAW["date2"] = day.split('_')[1]
        h_ADJUSTED["interval"] = h_RAW["interval"].copy()
        h_ADJUSTED["date"] = h_RAW["date"].copy()
        h_ADJUSTED["date2"] = h_RAW["date2"].copy()
        h_SSLA["interval"] = h_RAW["interval"].copy()
        h_SSLA["date"] = h_RAW["date"].copy()
        h_SSLA["date2"] = h_RAW["date2"].copy()
        h_CAPTIVE["interval"] = h_RAW["interval"].copy()
        h_CAPTIVE["date"] = h_RAW["date"].copy()
        h_CAPTIVE["date2"] = h_RAW["date2"].copy()
        for sheet in sheetnames:

            h_RAW[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 1::6].values.sum(axis= 1)
            h_ADJUSTED[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 2::6].values.sum(axis= 1)
            h_SSLA[sheet] = pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 3::6].values.sum(axis= 1)
            h_CAPTIVE[sheet]=pd.read_excel(day, skiprows = 11, sheet_name =sheet).iloc[:, 6::6].values.sum(axis= 1)
        h_RAW2= h_RAW2.append(h_RAW)
        h_SSLA2 = h_SSLA2.append(h_SSLA)
        h_ADJUSTED2 =h_ADJUSTED2.append(h_ADJUSTED)
        h_CAPTIVE2 = h_CAPTIVE2.append(h_CAPTIVE)

        print(day,":",len(pd.ExcelFile(day).sheet_names))
    
    #hourly to csv
            
    RAW2.to_csv('RAW'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    SSLA2.to_csv('SSLA'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    ADJUSTED2.to_csv('ADJUSTED'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    CAPTIVE2.to_csv('CAPTIVE'+'_'+RAW2.iloc[0,1]+'_'+RAW2.iloc[-1,1]+'_'+name+'.csv')
    print("five minute csv files created.")
        
        
        
        
    #fivemin to csv    
    h_RAW2.to_csv('RAW'+'_'+h_RAW2.iloc[0,1]+'_'+h_RAW2.iloc[-1,1]+'_'+h_name+'.csv')
    h_SSLA2.to_csv('SSLA'+'_'+h_RAW2.iloc[0,1]+'_'+h_RAW2.iloc[-1,1]+'_'+h_name+'.csv')
    h_ADJUSTED2.to_csv('ADJUSTED'+'_'+h_RAW2.iloc[0,1]+'_'+h_RAW2.iloc[-1,1]+'_'+h_name+'.csv')
    h_CAPTIVE2.to_csv('CAPTIVE'+'_'+h_RAW2.iloc[0,1]+'_'+h_RAW2.iloc[-1,1]+'_'+h_name+'.csv')
    print("hourly csv files created.")


# In[ ]:


print("enter any key to end script")
z = input()

