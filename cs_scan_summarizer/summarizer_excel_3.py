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
vendor_name = 'ABCCC'
analyst_name = 'DN'
date_batch = '20210801'
#######################

config_coles = r"config.json"
config_coles_clean = r"config2.json"

file_sql_summ = r"summarizer.sql"
file_sql_summ_vendor = r"summarizer_vendor.sql"
file_sql_cd_ref = r"cd_ref.sql"
file_sql_dept = r"dept.sql"
file_sql_gst = r"gst.sql"
# file_sql_cd_ref_listagg = r"cd_ref_listagg.sql"
# file_sql_cd_ref_listagg_item = r"cd_ref_listagg_item.sql"
file_sql_pct = r"count_pct.sql"

current_dir = 'D:\\python\\cs_scan_summarizer'
os.chdir('D:\\python\\cs_scan_summarizer')

path_excel = r"CS_SCAN_Vendorname_Analyst_Date.xlsx"
path_import_item = 'item_import.xlsx'
# vendor_name = (input('Input vendor name : ')).upper()
# analyst_name = (input('Input analyst name. Example: CT. Your answer is ')).upper()
# date_batch = input('Input date batch. Example: 20230207. Your answer is ')
iconPath_email = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
iconPath_excel = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"


path_export_final = 'CS_SCAN_'+vendor_name+'_'+analyst_name+'_'+date_batch+'.xlsx'
path_export_final_xlsb = 'CS_SCAN_'+vendor_name+'_'+analyst_name+'_'+date_batch+'.xlsb'
path_vba = 'CS_SCAN_vendorname_analyst_yyyymmdd.xlsb'

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

def item_gst(i):
    df = pd.read_excel(path_import_item,sheet_name=str(i))
    df['ITEM_IDNT'] = df['ITEM_IDNT'].astype(str)
    df['ITEM_IDNT'] = df['ITEM_IDNT'].str.strip() 
    item_unique = df['ITEM_IDNT'].drop_duplicates().tolist()
    item_unique = "','".join(item_unique)
    df['STATE'] = df['STATE'].str.strip()
    clm_start = df['CLM_START'][0]
    clm_end = df['CLM_END'][0]
    supp_number_filter = df['RMS_NUM'][0]
    print('item_unique',item_unique,'clm_end',clm_end)
    df_gst = connect_sql(cursor,file_sql= file_sql_gst, var_1=item_unique ,var_2 = clm_end)
    if np.isnan(supp_number_filter):
        supp_num = df_gst['SUPP_IDNT'][0]
    else:
        supp_num = supp_number_filter
        df_gst = df_gst[df_gst['SUPP_IDNT'] == str(supp_number_filter)].reset_index(drop=True)
    gst = df_gst['CML_COST_GST_RATE_PCT'][0]
    gst = int(gst)
    claim_number = f'{i}_{gst}'
    dept = df_gst['DEPT_IDNT'][0]
    supp_desc = df_gst['SUPP_DESC'][0]
    vendor_num = df_gst['VENDOR_NUM'][0]
    classify_state = df['CLASSIFY_STATE'][0]
    classify_promo_type = df['CLASSIFY_PROMO_TYPE'][0]
    pct = df['PERCENTAGE'][0]
    file_path_excel = df['EXCEL_PATH'][0]
    file_path_email = df['EMAIL_PATH'][0]
    if np.isnan(pct) == False:
        df_pct = connect_sql(cursor, file_sql_pct, var_1=item_unique, var_2= clm_start , var_3= clm_end )
        df = df.merge( right= df_pct , how = 'left',on ='ITEM_IDNT')
        df['RRP'] = (100- df['PERCENTAGE'].astype('float') )/ 100 * df['NORMAL_PRICE'].astype('float')
    else:
        pass
    if classify_state.lower() == 'state':
        # item_list_dict = df.set_index(['ITEM_IDNT','STATE'])[['RRP','SCANRATE']].to_dict('index')
        df = df[['ITEM_IDNT','STATE','RRP','SCANRATE']].drop_duplicates()
        df['SCANRATE'] = df['SCANRATE'].round(2)
        df['RRP'] = df['RRP'].round(2)
        df = df.groupby(by = ['STATE','RRP','SCANRATE'])['ITEM_IDNT'].agg(list).to_frame().reset_index()
        df['ITEM_IDNT'] = df['ITEM_IDNT'].apply(lambda x : convert_to_input_function(x))
        df = df.groupby(by = ['RRP','SCANRATE','ITEM_IDNT'])['STATE'].agg(list).to_frame().reset_index()
        df['STATE'] = df['STATE'].apply(lambda x : convert_to_input_sql(x))
        item_list_dict = df.set_index(['ITEM_IDNT','STATE'])[['RRP','SCANRATE']].to_dict('index')
        for key,value in item_list_dict.items():
            item_list_dict[key] = [item_list_dict[key]['RRP']] + [item_list_dict[key]['SCANRATE']] 
        print(df)
        print(item_list_dict)
    else:
        # item_list_dict = df.set_index('ITEM_IDNT')[['RRP','SCANRATE']].to_dict('index')
        df = df[['ITEM_IDNT','RRP','SCANRATE']].drop_duplicates()
        df['SCANRATE'] = df['SCANRATE'].round(2)
        df['RRP'] = df['RRP'].round(2)
        df = df.groupby(by = ['RRP','SCANRATE'])['ITEM_IDNT'].agg(list).to_frame().reset_index()
        df['ITEM_IDNT'] = df['ITEM_IDNT'].apply(lambda x : convert_to_input_function(x))
        item_list_dict = df.set_index('ITEM_IDNT')[['RRP','SCANRATE']].to_dict('index')
        for key,value in item_list_dict.items():
            item_list_dict[key] = [item_list_dict[key]['RRP']] + [item_list_dict[key]['SCANRATE']] 
        print(df)
        print(item_list_dict)
    for key,value in item_list_dict.items():
        if gst == 10:
            item_list_dict[key][0] = item_list_dict[key][0] /1.1 
        else:
            pass   
    return supp_num,supp_desc,vendor_num,supp_number_filter,claim_number,gst,clm_start,clm_end,dept,item_unique,item_list_dict,classify_state,file_path_excel,file_path_email,classify_promo_type

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

def create_worksheet(index_promo,gst,path_export_final):
    # Open Excel in background
    with xw.App(visible=False) as app:
        if index_promo == 1:
            wb_from = app.books.open(path_excel)
        else :
            wb_from = app.books.open(path_export_final)
        ws_from = wb_from.sheets['template']
        ws_from.copy(before=ws_from, name=str(index_promo)+'_'+str(gst))
        wb_from.save(path_export_final)
    return 'Done create worksheet'     

def remove_sheet_change_xlsb(sheet_name,path_export_final,path_export_final_xlsb):
    print('Start delete sheet & change to xlsb')
    with xw.App(visible=False) as app:
        wb = app.books.open(path_export_final)                
        wb.sheets[sheet_name].delete()
        wb.save(path_export_final_xlsb)
    try:
        os.remove(path_export_final)
    except Exception as e:
        print(e)
    return print('Done delete sheet & change to xlsb')


def df_sales_data(item_list_dict_gsted,classify_state,supp_number_filter,classify_promo_type):
    i = 0
    if classify_promo_type != 'TH':
        classify_promo_type = 'V2'
    print(classify_state,classify_promo_type)
    if classify_state.lower() == 'state':
        for key,value in item_list_dict_gsted.items():
            print(key,value)
            item_code,state = key
            if np.isnan(supp_number_filter):
                df_each_item = connect_sql(cursor,file_sql = file_sql_summ ,var_1 = classify_promo_type ,var_2 = item_code,var_3 = clm_start,var_4 =clm_end,var_5=value[0],var_6=value[1])
            else:
                df_each_item = connect_sql(cursor,file_sql = file_sql_summ_vendor,var_1 = classify_promo_type  ,var_2 = item_code,var_3 = clm_start,var_4 =clm_end,var_5=value[0],var_6 = supp_number_filter,var_7=value[1])
            state_filter = state.replace("'",'').split(',')
            df_each_item = df_each_item[df_each_item['RSTATE'].isin(state_filter)]
            if i == 0:
                df_merge = df_each_item
            else :
                df_merge = pd.concat([df_merge, df_each_item], ignore_index=True)
            i+=1
        df_merge['ELI_CLAIM'] = df_merge.RQTY_PROMO * df_merge.SCAN_RATE
        df_merge= df_merge.sort_values(by=['RSKU_ID','RDAY_DT','RSTATE'], ascending=True).reset_index(drop=True)
    else:
        for key,value in item_list_dict_gsted.items():
            print(key,value)
            if np.isnan(supp_number_filter):
                df_each_item = connect_sql(cursor,file_sql = file_sql_summ ,var_1 = classify_promo_type ,var_2 = key,var_3 = clm_start,var_4 =clm_end,var_5=value[0],var_6=value[1])
            else:
                df_each_item = connect_sql(cursor,file_sql = file_sql_summ_vendor ,var_1 = classify_promo_type  ,var_2 = key,var_3 = clm_start,var_4 =clm_end,var_5=value[0],var_6 = supp_number_filter,var_7=value[1])
            if i == 0:
                df_merge = df_each_item
            else :
                df_merge = pd.concat([df_merge, df_each_item], ignore_index=True)
            i+=1
        df_merge['ELI_CLAIM'] = df_merge.RQTY_PROMO * df_merge.SCAN_RATE
        df_merge= df_merge.sort_values(by=['RSKU_ID','RDAY_DT','RSTATE'], ascending=True).reset_index(drop=True)
    return df_merge

def product_state_summary(df_sales,df_state_ref):
    print('Start product_state_summary')
    list_data = []
    list_remove = []
    # Find distict var_1 and state
    # writer_excel(df,cell_export,length_start,count_df,length_end,number_sheet,path_export_final)
    df_temp =df_sales.drop_duplicates(['RSKU_ID','RITEM_DESC','RSTATE'])[['RSKU_ID','RITEM_DESC','RSTATE']]
    df_temp_2 = pd.merge(df_temp,df_state_ref,left_on=['RSKU_ID','RSTATE'],right_on=['ITEM_IDNT','CLM_STATE'], how='left')
    # print(df_ref)
    df_final = df_temp_2[['RSKU_ID','RITEM_DESC','RSTATE','REF_NUM','CLM_QTY','CLM_RATE','CLM_PRODUCT']]
    df_sku_desc = df_final[['RSKU_ID','RITEM_DESC']]
    df_state = df_final[['RSTATE']]
    df_ref = df_final[['REF_NUM','CLM_QTY','CLM_RATE','CLM_PRODUCT']]
    df_ref.insert(1,"REF_DESC",'')
    # Calculate number of rows
    number_rows_state = len(df_ref)
    dict_data_sku = {'df':df_sku_desc,'cell_export':'B121'}
    dict_data_state = {'df':df_state,'cell_export':'E121'}
    dict_data_remove = {'df':df_ref,'cell_export':'M121'}
    dict_remove = {'count_df':number_rows_state,'length_start':121,'length_end':601}
    list_data.append(dict_data_sku)
    list_data.append(dict_data_state)
    list_data.append(dict_data_remove)
    list_remove.append(dict_remove)
    print('Done product_state_summary')
    return list_data,list_remove


def product_summary(df_sales,df_item_ref):
    print('Start product_summary')
    list_data = []
    list_remove = []
    df_product =df_sales.drop_duplicates(['RSKU_ID','RITEM_DESC'])[['RSKU_ID','RITEM_DESC']]
    df_temp = pd.merge(df_product,df_item_ref,left_on=['RSKU_ID'],right_on=['ITEM_IDNT'], how='left')
    # df_final = df_temp[['RSKU_ID','RITEM_DESC','REF_NUM']]
    df_product_1 = df_temp[['RSKU_ID','RITEM_DESC']]
    df_ref_1 = df_temp[['REF_NUM','CLM_QTY','CLM_RATE','CLM_PRODUCT']]
    number_rows_sales = len(df_product)
    # writer_excel(df = df_product,path_export_final = path_export_final, cell_export = 'B20',number_sheet = number_sheet,length_start=20 , count_df=number_rows_sales, length_end=116)
    dict_data_sku = {'df':df_product_1,'cell_export':'B20'}
    dict_data_ref = {'df':df_ref_1,'cell_export':'L20'}
    dict_remove = {'count_df':number_rows_sales,'length_start':20,'length_end':116}
    list_data.append(dict_data_sku)
    list_data.append(dict_data_ref)
    list_remove.append(dict_remove)
    print('Done product_summary')
    return list_data , list_remove

def cd_ref(df_sales,supp_number_filter):
    print('Start cd ref')
    list_data = []
    list_remove = []
    df_ref = connect_sql(cursor,file_sql_cd_ref ,var_1 = item_unique, var_2= clm_start , var_3= clm_end, var_4 = clm_start, var_5 = clm_end)
    if np.isnan(supp_number_filter):
        pass
    else:
        df_ref = df_ref[df_ref['CLM_SUPPLIER_MERCH'] == str(supp_number_filter)]
    df_ref_groupby = df_ref.groupby('CLM_REF_NUM').agg({'CLM_PRODUCT':'sum'}).sort_values(by='CLM_PRODUCT', ascending=True).reset_index()
    df_sales_daily = pd.concat([df_sales, df_ref_groupby], axis=1 )
    print('Done cd ref')
    print('start state ref')
    df_state_ref = state_ref_groupby(df_ref)
    print('done state ref')
    print('start item ref')
    df_item_ref = item_ref_groupby(df_ref)
    print('done item ref')
    # writer_excel(df = df_sales, cell_export = 'B174',number_sheet= str(index_promo)+'_'+str(gst),length_start=174 ,count_df=len(df_sales), length_end=10174,path_export_final=path_export_final)
    dict_data = {'df':df_sales_daily,'cell_export':'B606'}
    dict_remove = {'count_df':len(df_sales),'length_start':606,'length_end':20606}
    list_data.append(dict_data)
    list_remove.append(dict_remove)
    return df_item_ref,df_state_ref,df_ref,df_sales_daily,list_data,list_remove

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

def move_worksheet_to_vba_template(path_xlsb):
    print('start move sheets')
    with xw.App(visible=False) as app:
        wb1 = app.books.open(path_xlsb)
        wb2 = app.books.open(path_vba)
        print(wb1.sheet_names)
        for sheet_name in wb1.sheet_names:
            ws1 = wb1.sheets(sheet_name)
            ws1.api.Copy(Before=wb2.sheets('Sheet1').api)
        wb2.sheets['Sheet1'].delete()
        wb1.close()
        wb2.save(path_xlsb)
    print('end move sheets')
    return None

def list_to_listagg(x):
    x = list(set(x))
    x_convert = ','.join(x)
    return x_convert

def state_ref_groupby(df_ref):
    df_ref_state_groupby = df_ref.groupby(by = ['ITEM_IDNT', 'CLM_STATE']).agg({'CLM_QTY':'sum','CLM_PRODUCT':'sum','CLM_REF_NUM':list}).reset_index()
    df_ref_state_groupby['REF_NUM'] = df_ref_state_groupby['CLM_REF_NUM'].apply(lambda x : list_to_listagg(x))
    df_ref_state_groupby['CLM_QTY'] = np.where(df_ref_state_groupby['CLM_QTY'].astype(int) != 0,df_ref_state_groupby['CLM_QTY'],np.nan) 
    df_ref_state_groupby['CLM_RATE'] = (df_ref_state_groupby['CLM_PRODUCT'] / df_ref_state_groupby['CLM_QTY']).astype('float').round(2)
    df_ref_state_groupby = df_ref_state_groupby[['ITEM_IDNT', 'CLM_STATE', 'CLM_QTY', 'CLM_RATE', 'CLM_PRODUCT', 'REF_NUM']]
    return df_ref_state_groupby

def item_ref_groupby(df_ref):
    df_ref_item_groupby = df_ref.groupby(by = ['ITEM_IDNT']).agg({'CLM_QTY':'sum','CLM_PRODUCT':'sum','CLM_REF_NUM':list}).reset_index()
    df_ref_item_groupby['REF_NUM'] = df_ref_item_groupby['CLM_REF_NUM'].apply(lambda x : list_to_listagg(x))
    df_ref_item_groupby['CLM_QTY'] = np.where(df_ref_item_groupby['CLM_QTY'].astype(int) != 0,df_ref_item_groupby['CLM_QTY'],np.nan) 
    df_ref_item_groupby['CLM_RATE'] = (df_ref_item_groupby['CLM_PRODUCT'] / df_ref_item_groupby['CLM_QTY']).astype('float').round(2)
    df_ref_item_groupby = df_ref_item_groupby[['ITEM_IDNT', 'CLM_QTY', 'CLM_RATE', 'CLM_PRODUCT', 'REF_NUM']]
    return df_ref_item_groupby


#MAIN
print('START')
cursor = set_up(config = config_coles)
excel_file = pd.ExcelFile(path_import_item)
count_sheets_excel_file = len(excel_file.sheet_names)
# excel_file.close()
summary_index_list =[]
for i in range(1,count_sheets_excel_file+1):
    print(f'Sheet {i}')
    supp_num,supp_desc,vendor_num,supp_number_filter,claim_number,gst,clm_start,clm_end,dept,item_unique,item_list_dict,classify_state,file_path_excel,file_path_email,classify_promo_type = item_gst(i)
    df_sales = df_sales_data(item_list_dict,classify_state,supp_number_filter,classify_promo_type)
    df_item_ref,df_state_ref,df_ref,df_sales_daily,list_data_sales,list_remove_sales  = cd_ref(df_sales,supp_number_filter)
    if df_ref.empty:
        prmt_id = ''
        prmt_name = ''
    else:
        prmt_id = df_ref['PRMTN_COMP_IDNT'][0]
        prmt_name = df_ref['PRMTN_COMP_NAME'][0]
    try:
        df_ref_groupby = df_ref.groupby('CLM_REF_NUM').agg({'CLM_PRODUCT':'sum'}).sort_values(by='CLM_PRODUCT', ascending=True).reset_index()
        ref_num_list = ', '.join(df_ref_groupby['CLM_REF_NUM'].tolist())
    except:
        ref_num_list = ''
    dict_data_dept = {'df':dept,'cell_export':'F8'}
    dict_data_supp_num = {'df':supp_num,'cell_export':'E8'}
    dict_data_supp_desc = {'df':supp_desc,'cell_export':'C8'}
    dict_data_vendor_num = {'df':vendor_num,'cell_export':'D8'}
    dict_data_claim_number = {'df': claim_number,'cell_export':'B16'}
    dict_data_prmt_id = {'df':prmt_id,'cell_export':'B12'}
    dict_data_prmt_name = {'df':prmt_name,'cell_export':'C12'}
    dict_data_ref_num_list = {'df':ref_num_list,'cell_export':'D16'}
    list_data_state,list_remove_state = product_state_summary(df_sales,df_state_ref)
    list_data_product ,list_remove_product = product_summary(df_sales,df_item_ref)
    list_data = list_data_sales + list_data_state + list_data_product + [dict_data_dept] + [dict_data_prmt_id] + [dict_data_prmt_name] +  [dict_data_supp_num] + [dict_data_supp_desc] + [dict_data_vendor_num] + [dict_data_claim_number] + [dict_data_ref_num_list]
    list_remove = list_remove_sales + list_remove_state + list_remove_product
    create_worksheet(i,gst,path_export_final)
    writer_excel(list_data,list_remove,claim_number,path_export_final)
    try:
        insert_attachments(str(i)+'_'+str(gst),file_path_excel,file_path_email,path_export_final)
    except:
        pass
    summary_index_list.append(claim_number)
fill_summary_sheet(summary_index_list,path_export_final=path_export_final) 
remove_sheet_change_xlsb(sheet_name = 'template',path_export_final=path_export_final ,path_export_final_xlsb = path_export_final_xlsb)
move_worksheet_to_vba_template(path_export_final_xlsb)

# xl.Application.Quit()
print('END')