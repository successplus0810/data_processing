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

config_coles = r"config.json"

file_sql_summ = r"summarizer.sql"
file_sql_ref_num = r"summarizer_ref_num.sql"
file_sql_ref_num_groupbyitem = r"summarizer_ref_num_GROUPBYITEM.sql"

current_dir = 'D:\\python\\co_scan_summarizer'
os.chdir('D:\\python\\co_scan_summarizer')


def connect_sql(cursor,file_sql,var_1,var_2 = '',var_3='',var_4='',var_5 = '',var_6 = '',var_7 = ''):
    print((open(file_sql).read()).format(var_1,var_2,var_3,var_4,var_5,var_6,var_7))
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

def list_to_listagg(list_ref_num):
    list_ref_num = list(set(list_ref_num))
    x = list(filter(lambda x: x.strip() != '', list_ref_num))
    x_convert = ','.join(x)
    return x_convert

def df_sales_data(cursor,supplier, promo_cat):
    df_ref_num_groupby = pd.DataFrame(columns=[
                            'CLM_REF_NUM',
                            'CLAIM_QTY',
                            'CLAIM_AMT'])
    df_promo_cat = connect_sql(cursor= cursor, file_sql=file_sql_summ, var_1 = promo_cat[3], var_2 = promo_cat[0], var_3 = supplier, var_4 = promo_cat[1],var_5 = promo_cat[6]) 
    gst = int(df_promo_cat['GST_RATE'].drop_duplicates().reset_index(drop=True)[0])
    dept_desc = df_promo_cat['DEPT_DESC'].drop_duplicates().reset_index(drop=True)[0]
    supp_desc = df_promo_cat['SUPP_DESC'].drop_duplicates().reset_index(drop=True)[0]
    supp_desc=''.join(filter(lambda x: x.isdigit() or x.isalpha() or x==' ', supp_desc))
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

def summarize_data(i,cursor,supplier,promo_cat):
    promo_name,vendor_num,gst,dept_desc,supp_desc,list_ref_num,excel_path,outlook_path,list_data_sales,list_remove_sales,df_sales = df_sales_data(cursor,supplier,promo_cat)
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
    return list_data,list_remove,supp_desc,claim_number,gst,excel_path,outlook_path

