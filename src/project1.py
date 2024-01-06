import pandas as pd
import numpy as np
import sys
import time
from datetime import datetime

def create_pera_shhet (arg1,arg2):
    df = pd.read_excel(arg1, sheet_name=None,header=0)
    print(df)
    c_df = pd.DataFrame()
    for key,value in df.items():
        if key not in ['Sheet1']:
            c_df =c_df.append(value.iloc[:,2:5], ignore_index=True)
    c_df.iloc[:,1] = c_df.iloc[:,1].str.replace('&','-')
    c_df.set_axis(["A", "B", "C", "D", "E", "F"], axis="columns", inplace=True)

    years = c_df.B.unique()
    years_list =[]
    for year in years :
        if str(year) != 'nan':
            years_list.append(year)
            
    years_list.sort()
    sheets_df_dic={}
    for year in years_list:
        sheet_df = c_df.loc[c_df['B'] == year]
        sheet_df.reset_index(inplace = True) 
        sheet_df = sheet_df.iloc[:,[1,3]]
        sheets_df_dic[year] = sheet_df
    currentDateTime = datetime.now().strftime("%m-%d-%Y ")
    writer = pd.ExcelWriter(f"File {currentDateTime}.xlsx", engine="xlsxwriter")
    workbook = writer.book
    merge_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1,'bold' : True})
    total_format = workbook.add_format({
        'font_size': 12,
        'bold' : True
    })
        
    for idx,key in enumerate(sheets_df_dic):
        sheets_df_dic[key].to_excel(writer, sheet_name='Sheet2', startcol=idx*4,startrow=2 ,index=False, header=False)
        worksheet = writer.sheets['Sheet2']
        worksheet.merge_range(0, idx*4, 0, idx*4+2, key, merge_format)
        column_sum = sheets_df_dic[key].iloc[:,[1]].sum()
        column_list = list(sheets_df_dic[key].iloc[:,0])
        values_list =[]
        for column in column_list:
            if str(column)!='nan':
                values_list.append(column)
        row_no = len(values_list)+2    
            
        worksheet.write(row_no,idx*4, 'Total', total_format)
        worksheet.write(row_no,idx*4+1, column_sum, total_format)
        
    writer.save()
    writer.close()    

    df3 = pd.read_excel(f"File {currentDateTime}.xlsx")
    times = 10
    dict_new = {}
    for i in range(times):
        n = i*4
        current_df = df3.iloc[:, [ n,n+1 ]]
        dict_new[f'old_df{i+1}'] = current_df
        
    df4 = pd.read_excel(arg2)
    times = 10
    dict_new2 = {}
    for i in range(times):
        n = i*4
        current_df = df4.iloc[:, [ n,n+1 ]]
        dict_new2[f'old_df{i+1}'] = current_df


    new_df = pd.DataFrame()
    writer = pd.ExcelWriter(f"Difference {currentDateTime}.xlsx", engine="xlsxwriter")
    workbook = writer.book
    merge_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1,'bold' : True})
    idx = 0
    for key in dict_new:
        row_no = base(dict_new[key])
        row_no2 = base(dict_new2[key])
        v1 = dict_new[key]
        v2 = dict_new2[key]
        v1 = v1.drop(row_no)
        v2 = v2.drop(row_no2)
        final_df = pd.concat([v1, v2])
        final2 = final_df.drop_duplicates(keep=False)   
        final2.to_excel(writer, sheet_name='Sheet1', startcol=idx*2,startrow=0 ,index=False)
        worksheet.merge_range(0, idx*2, 0, idx*2+2, key, merge_format)
        idx = idx+1
        
        
    writer.save()
    writer.close() 



def base(dict1):
    column_list = list(dict1.iloc[:,0])
    values_list =[]
    for column in column_list:
        if str(column)!='nan':
            values_list.append(column)
    row_nu = len(values_list)
    return row_nu


if __name__ == "__main__" :
    arg1 = sys.argv[1]
    arg2 = sys.argv[2]
    create_pera_shhet(arg1,arg2)
      