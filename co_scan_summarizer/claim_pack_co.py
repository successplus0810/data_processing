import json
import pandas as pd
import snowflake.connector as sf
import os
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
import datetime
import numpy as np
import win32com.client as win32
from pywintypes import com_error
import math
import summarizer_ONLINE 
import summarizer_EXCLUSIVE

##########################' 
# Analyst fill
analyst_name = 'DN'
date_batch = '20210801'
month_filter = '202110'
#######################

config_coles = r"config.json"

file_sql_summ = r"summarizer.sql"
file_sql_ref_num = r"summarizer_ref_num.sql"
file_sql_ref_num_groupbyitem = r"summarizer_ref_num_GROUPBYITEM.sql"
file_sql_claimpack = r"claim_pack.sql"
file_sql_claimpack_schema = r"claim_pack_schema.sql"

current_dir = 'D:\\python\\co_scan_summarizer'
os.chdir('D:\\python\\co_scan_summarizer')

path_excel = r"CS_SCAN ONLINE_Vendorname_Analyst_Date.xlsx"
# path_import_item = 'item_import.xlsx'
path_check_list_promo = fr"D:\\python\\co_scan_summarizer\\claim_pack_co\\{month_filter}\\check_list_promo.xlsx"
path_export = fr"D:\\python\\co_scan_summarizer\\claim_pack_co\\{month_filter}\\"
iconPath_email = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
iconPath_excel = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"


def set_up(config):
    """Set up connection to SnowFlake"""
    config = json.loads(open(config).read())
    account = config['snowflake']['account']
    user = config['snowflake']['user']
    warehouse = config['snowflake']['warehouse']
    role = config['snowflake']['role']
    database = config['snowflake']['database']
    schema = config['snowflake']['schema']
    password = config['snowflake']['password']
    auth = config['snowflake']['authenticator']

    conn = sf.connect(user=user, password=password, account=account, authenticator=auth,
                      warehouse=warehouse, role=role, database=database, schema=schema)

    cursor = conn.cursor()
    return cursor
def connect_sql(cursor,file_sql,var_1,var_2 = '',var_3='',var_4='',var_5 = '',var_6 = '',var_7 = ''):
    try:
        print((open(file_sql).read()).format(var_1,var_2,var_3,var_4,var_5,var_6,var_7))
        cursor.execute((open(file_sql).read()).format(var_1,var_2,var_3,var_4,var_5,var_6,var_7))
        all_rows = cursor.fetchall()
        field_names = [i[0] for i in cursor.description]
    finally:
        # conn.close()
        pass
    df = pd.DataFrame(all_rows)
    try:
        df.columns = field_names
    except ValueError:
        return pd.DataFrame(columns=field_names)
    return df


def writer_excel(data,remove,number_sheet,path_export_final):
    # data = list_data, remove = list_remove,number_sheet= str(index_promo)+'_'+str(gst),path_export_final=path_export_final
    #select sheet
    sheet_df_mapping = {number_sheet: data}
    sheet_df_remove  = {number_sheet: remove}
    # Open Excel in background
    with xw.App(visible=False) as app:
        wb = app.books.open(path_export_final)
        # List of current worksheet names
        current_sheets = [sheet.name for sheet in wb.sheets]
        # Iterate over sheet/df mapping
        # If sheet already exist, overwrite current cotent. Else, add new sheet
        print('start copy data')
        for sheet_name in sheet_df_mapping.keys():
            if sheet_name in current_sheets:
                for df_data in data :
                    wb.sheets(sheet_name).range(df_data['cell_export']).options(index=False,header=False).value = df_data['df']
            else:
                'Name of sheet cannot be found in Excel file, please check again'
        print('done copy data')
        print('start delete rows')
        for sheet_name in sheet_df_remove.keys():
            if sheet_name in current_sheets:
                for df_remove in remove :
                    # wb.sheets(sheet_name).range(df_cell['cell_export']).options(index=False,header=False).value = df_cell['df']
                    length_start = df_remove['length_start'] + df_remove['count_df']
                    range_length_to_remove = str(length_start)+':'+ str(df_remove['length_end'])
                    wb.sheets(sheet_name).range(range_length_to_remove).api.Delete(DeleteShiftDirection.xlShiftUp)
            else:
                'Name of sheet cannot be found in Excel file, please check again'
        print('done delete rows')
        wb.save(path_export_final)
    return None

def fill_summary_sheet(summary_index_list,path_export_final):
    print('Start fill summary sheet')
    with xw.App(visible=False) as app:
        wb_from = app.books.open(path_export_final)
        summary_index = 1
        for index in summary_index_list:
            wb_from.sheets['Vendor Summary'].range('B'+str(summary_index+10)).value = index
            summary_index += 1
        length_start = summary_index + 10
        range_length_to_remove = str(length_start)+':'+ str(30)
        print(range_length_to_remove)
        wb_from.sheets('Vendor Summary').range(range_length_to_remove).api.Delete(DeleteShiftDirection.xlShiftUp)         
        wb_from.save(path_export_final)
    return 'Done fill summary sheet' 

def create_worksheet(index_promo,gst,path_export_final,template_name):
    # Open Excel in background
    with xw.App(visible=False) as app:
        if index_promo == 1:
            wb_from = app.books.open(path_excel)
        else :
            wb_from = app.books.open(path_export_final)
        ws_from = wb_from.sheets[template_name]
        ws_from.copy(before=ws_from, name=str(index_promo)+'_'+str(gst))
        wb_from.save(path_export_final)
    return 'Done create worksheet'     

def remove_sheet(path_export_final):
    print('Start delete sheet ')
    with xw.App(visible=False) as app:
        wb = app.books.open(path_export_final)                
        wb.sheets['template'].delete()
        wb.sheets['template_nostate'].delete()
        wb.save(path_export_final)
    return print('Done delete sheet & change to xlsb')


def insert_attachments(sheet_name,file_path_excel,file_path_email,path_export_final):  
    print('Start insert email and excel')
    print(file_path_excel)
    print(file_path_email)
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    wb = xl.Workbooks.Open(fr'{current_dir}\{path_export_final}', UpdateLinks = True)
    ws = wb.Worksheets(sheet_name)
    try:
        excel_name = file_path_excel.split('/')[-1][0:20]
    except:
        excel_name = ''
    try:
        email_name = file_path_email.split('/')[-1][0:20]
    except:
        email_name = ''
    obj = ws.OLEObjects()
    xl.DisplayAlerts = False
    #xl.AskToUpdateLinks = False
    try:
        obj.Add(ClassType=None, Filename=file_path_excel, Link=False, DisplayAsIcon=True, IconFileName=iconPath_excel,IconIndex=0, IconLabel = excel_name , Left=ws.Range("J8").Left, Top=ws.Range("J8").Top, Width=50, Height=50)
        print(f'Successfully insert excel file in sheet {sheet_name}')
    except com_error:
        print(f'Cannot insert excel file in sheet {sheet_name}')
        pass
    try:
        obj.Add(ClassType=None, Filename=file_path_email, Link=False, DisplayAsIcon=True, IconFileName=iconPath_email,IconIndex=0, IconLabel = email_name , Left= ws.Range("L8").Left, Top=ws.Range("L8").Top, Width=50, Height=50)
        print(f'Successfully insert email file in sheet {sheet_name}')
    except com_error:
        print(f'Cannot insert email file in sheet {sheet_name}')
        pass
    xl.DisplayAlerts = True
    #xl.AskToUpdateLinks = True
    wb.Save()
    wb.Close()
    # xl.Application.Quit()
    #del xl
    print('Done insert email and excel')
    return None

def create_folder_if_not_exists(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print("Folder created: " + folder_path)
    else:
        print("Folder already exists: ", folder_path)



#MAIN
print('START')

create_folder_if_not_exists(fr"D:\python\co_scan_summarizer\claim_pack_co\{month_filter}")

cursor = set_up(config = config_coles)
df_ONLINE= connect_sql(cursor = cursor, file_sql= file_sql_claimpack_schema, var_1= 'ELIGIBLE_2', var_2= 'VIEW_MULTIBUYPROMOS_COLOTHER_SCHEMA', var_3 = month_filter)
df_MULTIBUYS= connect_sql(cursor = cursor, file_sql= file_sql_claimpack_schema, var_1= 'ELIGIBLE', var_2= 'VIEW_CO_EXC_TH_SCHEMA', var_3 = month_filter)
df_SIMPLE= connect_sql(cursor = cursor, file_sql= file_sql_claimpack_schema, var_1= 'ELIGIBLE', var_2= 'VIEW_CO_EXC_SIMP_SCHEMA', var_3 = month_filter)

df_ONLINE.to_csv(path_or_buf= fr"{path_export}checklist_ONLINE.csv", index= False)
df_MULTIBUYS.to_csv(path_or_buf= fr"{path_export}checklist_MULTIBUYS.csv", index= False)
df_SIMPLE.to_csv(path_or_buf= fr"{path_export}checklist_SIMPLE.csv", index= False)

df_excel = connect_sql(cursor = cursor, file_sql= file_sql_claimpack, var_1= month_filter)
df_excel = df_excel[['PROMO_ID','SUPPLIER_ID','CAT_NUM','CLASSIFY_NOTE']].applymap(str)
i = 1
dict_supplier = {}
dict_item_import = df_excel.to_dict(orient='records')
for index in dict_item_import:
    if index['CLASSIFY_NOTE'].upper() == 'ONLINE':
        tbl_daily = 'COLES.STI_WIP_CO.VIEW_MULTIBUYPROMOS_COLOTHER_DAILY'
        tbl_promo = 'COLES.STI_WIP_CO.VIEW_MULTIBUYPROMOS_COLOTHER_SCHEMA'
        column_list =  'ORDER_STAGING_DAY_IDNT, ITEM_IDNT_VCHAR ITEM_IDNT, ITEM_LONG_DESC ITEM_NAME, STATE CLM_STATE, PICKED_QTY, SCAN_TTL, REF_NUM CLM_REF_NUM, CLAIM_QTY, CLAIM_AMT, ELIGIBLE, PRMTN_COMP_IDNT, PRMTN_COMP_NAME, PROMO_START_DT, PROMO_END_DT, GST_RATE, PAF_LOCATION, EMAIL, SUPP_IDNT, SUPP_DESC, DEPT_IDNT, DEPT_DESC, SAP_ID'
        template = 'template'
    elif index['CLASSIFY_NOTE'].upper() == 'ONLINE SIMPLE EXCLUSIVE': # simple exclusive
        tbl_daily = '(SELECT *, NULL AS STATE FROM COLES.STI_WIP_CO.VIEW_CO_EXC_SIMP_DAILY)'
        tbl_promo = 'COLES.STI_WIP_CO.VIEW_CO_EXC_SIMP_SCHEMA'
        column_list = 'ORDER_STAGING_DAY_IDNT, ITEM_IDNT_VCHAR ITEM_IDNT, ITEM_LONG_DESC ITEM_NAME, STATE CLM_STATE, PICKED_QTY, SCAN_TTL, PRMTN_COMP_IDNT, PRMTN_COMP_NAME, PROMO_START_DT, PROMO_END_DT, GST_RATE, PAF_LOCATION, EMAIL, SUPP_IDNT, SUPP_DESC, DEPT_IDNT, DEPT_DESC, SAP_ID'
        template = 'template_nostate'
    else: #multibuy exclusive
        tbl_daily = 'COLES.STI_WIP_CO.VIEW_CO_EXC_TH_DAILY'
        tbl_promo = 'COLES.STI_WIP_CO.VIEW_CO_EXC_TH_SCHEMA'
        column_list = 'ORDER_STAGING_DAY_IDNT, ITEM_IDNT_VCHAR ITEM_IDNT, ITEM_LONG_DESC ITEM_NAME, STATE CLM_STATE, PICKED_QTY, SCAN_TTL, PRMTN_COMP_IDNT, PRMTN_COMP_NAME, PROMO_START_DT, PROMO_END_DT, GST_RATE, PAF_LOCATION, EMAIL, SUPP_IDNT, SUPP_DESC, DEPT_IDNT, DEPT_DESC, SAP_ID'
        template = 'template_nostate'
    if index['SUPPLIER_ID'] not in dict_supplier.keys():
        dict_supplier[index['SUPPLIER_ID']] = []
        dict_supplier[index['SUPPLIER_ID']].append([index['PROMO_ID'],index['CAT_NUM'],index['CLASSIFY_NOTE'],tbl_daily,template,tbl_promo,column_list])
    else:
        dict_supplier[index['SUPPLIER_ID']].append([index['PROMO_ID'],index['CAT_NUM'],index['CLASSIFY_NOTE'],tbl_daily,template,tbl_promo,column_list])

print(dict_supplier)
check_list_promo_index = 1
for supplier, list_promo_cat in dict_supplier.items():
    summary_index_list =[]
    i = 1
    for promo_cat in list_promo_cat:
        # break
        if promo_cat[2] == 'ONLINE':
            list_data,list_remove,supp_desc,claim_number,gst,excel_path,outlook_path = summarizer_ONLINE.summarize_data(i = i, cursor = cursor, supplier= supplier , promo_cat= promo_cat)
        # elif promo_cat[2] == 'ONLINE MULTIBUYS EXCLUSIVE':
        else:
            list_data,list_remove,supp_desc,claim_number,gst,excel_path,outlook_path = summarizer_EXCLUSIVE.summarize_data(i = i, cursor = cursor, supplier= supplier , promo_cat= promo_cat)
        supp_desc=''.join(filter(lambda x: x.isdigit() or x.isalpha() or x==' ', supp_desc))
        df_sales = list_data[0]['df']
        df_sales['amount'] = df_sales['SCAN_TTL'].astype(float) * df_sales['PICKED_QTY'].astype(float)
        sale_amount = df_sales['amount'].astype(float).sum() - df_sales['CLAIM_AMT'].astype(float).sum()
        #path export
        path_export_final = path_export + 'CS_SCAN ONLINE_'+supp_desc+'_'+analyst_name+'_'+date_batch+ '_' +supplier +'.xlsx'
        template_name = promo_cat[4]
        create_worksheet(i,gst,path_export_final,template_name)
        writer_excel(list_data,list_remove,claim_number,path_export_final)
        try:
            insert_attachments(str(i)+'_'+str(gst),excel_path,outlook_path,path_export_final)
        except:
            pass

        check_list_promo_index += 1
        with xw.App(visible=False) as app:
            print('check_list_promo_index')
            if check_list_promo_index == 2:
                wb = app.books.open('check_list_promo.xlsx')
            else:
                wb = app.books.open(path_check_list_promo)
            wb_sheet = wb.sheets['Sheet1']
            check_list_promo = [supplier] + [supp_desc] +[promo_cat[0]] + [promo_cat[1]] + [promo_cat[2]] +[sale_amount] +['Done'] 
            print(check_list_promo)
            wb_sheet.range(f'A{check_list_promo_index}').value =  check_list_promo
            wb.save(path_check_list_promo)  
        
        summary_index_list.append(claim_number)
        #Checklist promo
        i+=1
    fill_summary_sheet(summary_index_list,path_export_final=path_export_final) 
    remove_sheet(path_export_final=path_export_final)
print('----------END---------------')