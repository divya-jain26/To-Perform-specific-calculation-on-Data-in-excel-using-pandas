# -*- coding: utf-8 -*-
"""
Created on Tue Jan 18 12:55:53 2022

@author: ejaidiv
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Jan 14 21:41:20 2022

@author: ejaidiv
"""

import pandas as pds
newData = pds.read_excel('C:\OOS\OOS_Data.xlsx','Data1',skiprows=1)
Category = '';

data = []
for index, row in newData.iterrows():
    if not (pds.isna(row[0])):
        if(row[0] == 'New_data_cell_a_list'):
            technology = 'A'
        elif(row[0] == 'New_data_cell_b_list'):
            technology = 'B'
        elif(row[0] == 'New_data_cell_c_list'):
            technology = 'C'
        
            

        if not pds.isna(row[3]):
            newData.at[index, 9] = Category


newData.dropna(how='any', inplace=True)




newData1 = pds.read_excel('C:\OOS\OOS_Data.xlsx','Data2',skiprows=1)
Category = '';


for index, row in newData1.iterrows():
    if not (pds.isna(row[0])):
        if(row[0] == 'New_data_cell_a_list'):
            technology = 'A'
        elif(row[0] == 'New_data_cell_b_listt'):
            technology = 'B'
        elif(row[0] == 'New_data_cell_c_list'):
            technology = 'C'
        
            

        if not pds.isna(row[3]):
            newData1.at[index, 9] = Category
            

               
newData1.dropna(how='any', inplace=True)
pds.set_option('display.max_rows', 900)

newData2=pds.concat([newData,newData1])


newData2.iloc[:,0]=newData2.iloc[:,0].str[-6:]
newData2.columns = ["Name","class","ID","City","State","Code","Conatact no","Name_ID","Locker_id","Category"]


df6=pds.read_excel('C:\OOS\OOS_Data.xlsx','Ref')
df6['Code'] = df6['Code'].str.upper()


final_data = pds.merge(newData2,df6, on='Code',how ='left')

final_data=final_data.drop(columns=['Flag','MAGNET','PIn ID','Status','Number','ID'])

final_data=pds.DataFrame(final_data)

df = final_data[final_data.Status == 0]

#abc=final_data.groupby(['Tech','Region'])['Region'].aggregate('count').unstack()
#abc = pds.pivot_table(final_data,values=Region.count(),index=['Tech'],columns=['Region'],aggfunc=sum)
#abc

abc=df.groupby(['City','State'])['State'].aggregate('count').unstack()
#abc = pds.pivot_table(final_data,values=Region.count(),index=['Tech'],columns=['Region'],aggfunc=sum)
abc

with pds.ExcelWriter('C:\OOS\OOS_output.xlsx') as writer:
    df.to_excel(writer, sheet_name="Final")
    abc.to_excel(writer, sheet_name="Table")
    