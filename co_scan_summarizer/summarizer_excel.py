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

##########################' 
# Analyst fill
analyst_name = 'DN'
date_batch = '20210801'
#######################

config_coles = r"config.json"

file_sql_summ = r"summarizer.sql"

current_dir = 'D:\\python\\co_scan_summarizer'
os.chdir('D:\\python\\co_scan_summarizer')

path_excel = r"CS_SCAN ONLINE_Vendorname_Analyst_Date.xlsx"
path_import_item = 'item_import.xlsx'
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
   
def convert_to_input_sql(num_list):
    num_list_final = ''
    # print('SUPP LIST',supp_num_list)
    for num_list in num_list:
        num_list_final = num_list_final + "'" + num_list + "',"
    return num_list_final[:-1]

def convert_to_input_function(num_list):
    num_list_final = ''
    # print('SUPP LIST',supp_num_list)
    for num_list in num_list:
        num_list_final = num_list_final + num_list + ','
    return num_list_final[:-1]


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
        wb.sheets['template_simple'].delete()
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

def list_to_listagg(list_ref_num):
    list_ref_num = list(set(list_ref_num))
    x = list(filter(lambda x: x.strip() != '', list_ref_num))
    x_convert = ','.join(x)
    return x_convert

def df_sales_data(supplier, promo_cat):
    df_ref_num_groupby = pd.DataFrame(columns=[
                            'CLM_REF_NUM',
                            'CLAIM_QTY',
                            'CLAIM_AMT'])
    df_promo_cat = connect_sql(cursor= cursor, file_sql=file_sql_summ, var_1 = promo_cat[3], var_2 = promo_cat[0], var_3 = supplier, var_4 = promo_cat[1]) 
    gst = int(df_promo_cat['GST_RATE'].drop_duplicates().reset_index(drop=True)[0])
    dept_desc = df_promo_cat['DEPT_DESC'].drop_duplicates().reset_index(drop=True)[0]
    supp_desc = df_promo_cat['SUPP_DESC'].drop_duplicates().reset_index(drop=True)[0]
    # ref_num = ', '.join(df_promo_cat['CLM_REF_NUM'].drop_duplicates().to_list())
    excel_path = df_promo_cat['PAF_LOCATION'].drop_duplicates().reset_index(drop=True)[0]
    outlook_path = df_promo_cat['EMAIL'].drop_duplicates().reset_index(drop=True)[0]
    vendor_num = df_promo_cat['SAP_ID'].drop_duplicates().reset_index(drop=True)[0]
    promo_name = df_promo_cat['PRMTN_COMP_NAME'].drop_duplicates().reset_index(drop=True)[0]
    # df_promo_cat.columns.to_list()
    df_sales = df_promo_cat[['ORDER_STAGING_DAY_IDNT',
                            'ITEM_NAME',
                            'ITEM_IDNT',
                            'CLM_STATE',
                            'PICKED_QTY',
                            'SCAN_TTL',
                            'CLM_REF_NUM',
                            'CLAIM_QTY',
                            'CLAIM_AMT',
                            'PRMTN_COMP_IDNT',
                            'PRMTN_COMP_NAME']].reset_index(drop=True)
    #group by ref_num
    list_ref_num = df_sales['CLM_REF_NUM'].drop_duplicates().reset_index(drop=True).to_list()
    if len(list_ref_num) == 1 and list_ref_num[0].strip() == '':
        print('no ref')
    else:
        list_ref_num = list(filter(lambda x: x.strip() != '', list_ref_num)) 
        df_ref_num = df_sales[['CLM_REF_NUM','CLAIM_QTY','CLAIM_AMT']]
        df_ref_num_groupby = df_ref_num.groupby(by = ['CLM_REF_NUM']).agg({'CLAIM_QTY':'sum','CLAIM_AMT':'sum'}).reset_index()
        df_ref_num_groupby = df_ref_num_groupby.query('CLM_REF_NUM in @list_ref_num')
    df_sales_concat = pd.concat([df_sales[['ORDER_STAGING_DAY_IDNT',
                        'ITEM_NAME',
                        'ITEM_IDNT',
                        'CLM_STATE',
                        'PICKED_QTY',
                        'SCAN_TTL',
                        'PRMTN_COMP_IDNT',
                        'PRMTN_COMP_NAME']].reset_index(drop=True), df_ref_num_groupby.reset_index(drop=True)], axis=1)
    df_sales_concat = df_sales_concat[['ORDER_STAGING_DAY_IDNT',
                                'ITEM_NAME',
                                'ITEM_IDNT',
                                'CLM_STATE',
                                'PICKED_QTY',
                                'SCAN_TTL',
                                'CLM_REF_NUM',
                                'CLAIM_QTY',
                                'CLAIM_AMT',
                                'PRMTN_COMP_IDNT',
                                'PRMTN_COMP_NAME']]
    ########
    list_data = []
    list_remove = []
    # if promo_cat[2] != 'ONLINE SIMPLE EXCLUSIVE':
    dict_data = {'df':df_sales_concat[['ORDER_STAGING_DAY_IDNT',
                                'ITEM_NAME',
                                'ITEM_IDNT',
                                'CLM_STATE',
                                'PICKED_QTY',
                                'SCAN_TTL',
                                'CLM_REF_NUM',
                                'CLAIM_QTY',
                                'CLAIM_AMT']],'cell_export':'B607'}
    dict_data_2 = {'df':df_sales_concat[[
                                'PRMTN_COMP_IDNT',
                                'PRMTN_COMP_NAME']],'cell_export':'L607'}
    dict_remove = {'count_df':len(df_sales_concat),'length_start':607,'length_end':20607}
    # else:
    #     dict_data = {'df': df_sales,'cell_export':'B121'}
    #     dict_data_2 = {'df':pd.DataFrame(),'cell_export':'B607'}
    #     dict_remove = {'count_df':len(df_sales),'length_start':121,'length_end':20121}
    list_data.append(dict_data)
    list_data.append(dict_data_2)
    list_remove.append(dict_remove)
    return promo_name,vendor_num,gst,dept_desc,supp_desc,list_ref_num,excel_path,outlook_path,list_data,list_remove,df_sales

def product_state_summary(df_sales):
    print('Start product_state_summary')
    list_data = []
    list_remove = []
    # Find distict var_1 and state
    # writer_excel(df,cell_export,length_start,count_df,length_end,number_sheet,path_export_final)
    df_state_groupby = df_sales.groupby(by = ['ITEM_IDNT', 'ITEM_NAME','CLM_STATE']).agg({'CLM_REF_NUM':list,'CLAIM_QTY':'sum','CLAIM_AMT':'sum'}).reset_index()
    df_state_groupby['CLM_REF_NUM'] = df_state_groupby['CLM_REF_NUM'].apply(lambda x : list_to_listagg(x))
    df_state_groupby = df_state_groupby[['ITEM_IDNT', 'ITEM_NAME','CLM_STATE', 'CLM_REF_NUM','CLAIM_QTY','CLAIM_AMT']]
    # Calculate number of rows
    dict_data_sku = {'df':df_state_groupby[['ITEM_IDNT', 'ITEM_NAME','CLM_STATE']],'cell_export':'B121'}
    dict_data_ref_num = {'df':df_state_groupby[['CLM_REF_NUM','CLAIM_QTY']],'cell_export':'G121'}
    dict_data_ref_num_amt = {'df':df_state_groupby[['CLAIM_AMT']],'cell_export':'J121'}
    dict_remove = {'count_df':len(df_state_groupby),'length_start':121,'length_end':602}
    list_data.append(dict_data_sku)
    list_data.append(dict_data_ref_num)
    list_data.append(dict_data_ref_num_amt)
    list_remove.append(dict_remove)
    print('Done product_state_summary')
    return list_data,list_remove


def product_summary(df_sales):
    print('Start product_summary')
    list_data = []
    list_remove = []
    # Find distict var_1 and state
    # writer_excel(df,cell_export,length_start,count_df,length_end,number_sheet,path_export_final)
    df_item_groupby = df_sales.groupby(by = ['ITEM_IDNT', 'ITEM_NAME']).agg({'CLM_REF_NUM':list,'CLAIM_QTY':'sum','CLAIM_AMT':'sum'}).reset_index()
    df_item_groupby['CLM_REF_NUM'] = df_item_groupby['CLM_REF_NUM'].apply(lambda x : list_to_listagg(x))
    df_item_groupby = df_item_groupby[['ITEM_IDNT','ITEM_NAME', 'CLM_REF_NUM','CLAIM_QTY','CLAIM_AMT']]
    # Calculate number of rows
    dict_data_sku = {'df':df_item_groupby[['ITEM_IDNT', 'ITEM_NAME']],'cell_export':'B20'}
    dict_data_ref_num = {'df':df_item_groupby[['CLM_REF_NUM','CLAIM_QTY']],'cell_export':'F20'}
    dict_data_ref_num_amt = {'df':df_item_groupby[['CLAIM_AMT']],'cell_export':'I20'}
    dict_remove = {'count_df':len(df_item_groupby),'length_start':20,'length_end':116}
    list_data.append(dict_data_sku)
    list_data.append(dict_data_ref_num)
    list_data.append(dict_data_ref_num_amt)
    list_remove.append(dict_remove)
    print('Done product_summary')
    return list_data,list_remove


#MAIN
print('START')
cursor = set_up(config = config_coles)
df_excel = pd.read_excel(path_import_item,sheet_name='1')
df_excel = df_excel[['PROMO_ID','SUPPLIER_ID','CAT_NUM','CLASSIFY_NOTE']].applymap(str)
dict_supplier = {}
dict_item_import = df_excel.to_dict(orient='records')
for index in dict_item_import:
    if index['CLASSIFY_NOTE'].upper() == 'ONLINE':
        tbl_daily = 'COLES.STI_WIP_CO.VIEW_MULTIBUYPROMOS_COLOTHER_DAILY'
        template = 'template'
    elif index['CLASSIFY_NOTE'].upper() == 'ONLINE SIMPLE EXCLUSIVE': # simple exclusive
        tbl_daily = 'COLES.STI_WIP_CO.VIEW_CO_EXC_SIMP_DAILY'
        template = 'template_simple'
    else: #multibuy exclusive
        tbl_daily = 'COLES.STI_WIP_CO.VIEW_CO_EXC_TH_DAILY'
        template = 'template'
    if index['SUPPLIER_ID'] not in dict_supplier.keys():
        dict_supplier[index['SUPPLIER_ID']] = []
        dict_supplier[index['SUPPLIER_ID']].append([index['PROMO_ID'],index['CAT_NUM'],index['CLASSIFY_NOTE'],tbl_daily,template])
    else:
        dict_supplier[index['SUPPLIER_ID']].append([index['PROMO_ID'],index['CAT_NUM'],index['CLASSIFY_NOTE'],tbl_daily,template])

for supplier, list_promo_cat in dict_supplier.items():
    summary_index_list =[]
    i = 1
    for promo_cat in list_promo_cat:
        promo_name,vendor_num,gst,dept_desc,supp_desc,list_ref_num,excel_path,outlook_path,list_data_sales,list_remove_sales,df_sales = df_sales_data(supplier,promo_cat)
        supp_desc=''.join(filter(lambda x: x.isdigit() or x.isalpha() or x==' ', supp_desc))
        list_data_state,list_remove_state = product_state_summary(df_sales)
        list_data_item,list_remove_item = product_summary(df_sales)
        claim_number = f'{i}_{gst}'
        dict_data_dept = {'df':promo_cat[1],'cell_export':'F8'}
        dict_data_supp_num = {'df':supplier,'cell_export':'E8'}
        dict_data_supp_desc = {'df':supp_desc,'cell_export':'C8'}
        dict_data_vendor_num = {'df':vendor_num,'cell_export':'D8'}
        dict_data_claim_number = {'df': claim_number,'cell_export':'B16'}
        dict_data_prmt_id = {'df':promo_cat[0],'cell_export':'B12'}
        dict_data_prmt_name = {'df':promo_name,'cell_export':'C12'}
        dict_data_ref_num_list = {'df':list_to_listagg(list_ref_num),'cell_export':'D16'}
        dict_data_ref_num_list = {'df':promo_cat[2],'cell_export':'N11'}
        list_data = list_data_sales + list_data_state + list_data_item + [dict_data_dept] + [dict_data_prmt_id] + [dict_data_prmt_name] +  [dict_data_supp_num] + [dict_data_supp_desc] + [dict_data_vendor_num] + [dict_data_claim_number] + [dict_data_ref_num_list]
        list_remove = list_remove_sales + list_remove_state + list_remove_item
        #path export
        path_export_final = 'CS_SCAN ONLINE_'+supp_desc+'_'+analyst_name+'_'+date_batch+'.xlsx'
        print(f'-------------------{path_export_final}-------------------')
        template_name = promo_cat[4]
        create_worksheet(i,gst,path_export_final,template_name)
        writer_excel(list_data,list_remove,claim_number,path_export_final)
        try:
            insert_attachments(str(i)+'_'+str(gst),excel_path,outlook_path,path_export_final)
        except:
            pass
        summary_index_list.append(claim_number)
        i+=1
    fill_summary_sheet(summary_index_list,path_export_final=path_export_final) 
    remove_sheet(path_export_final=path_export_final)

print('----------END---------------')