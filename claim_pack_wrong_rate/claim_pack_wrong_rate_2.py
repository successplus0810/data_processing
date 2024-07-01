import json 
import pandas as pd
import snowflake.connector as sf
import os
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
import datetime
from pywintypes import com_error
import win32com.client as win32
import logging
import math
pd.options.mode.chained_assignment = None

# current_dir = os.getcwd()
# os.chdir(r'D:\python\claim_pack_python')
os.chdir("D:\\python\\claim_pack_wrong_rate")
current_dir = r'D:\python\claim_pack_wrong_rate'
###### Analyst fill
folder_name = '202204'
month_filter = '202204'
path_check_list_xlsx = r'checklist_xlsx.xlsx'
###############################################
config_coles = r"config.json"

file_sql_claimpack = r"claim_pack_state.sql"
file_sql_summ = r"summarizer.sql"
file_sql_summ_state = r"summarizer_state.sql"
file_sql_summ_th = r"summarizer_th.sql"
file_sql_summ_state_th = r"summarizer_state_th.sql"
file_sql_cd_ref = r"cd_ref.sql"
file_sql_cd_check_cole_online = r"cd_check_cole_online.sql"
file_sql_ven_stop_trading = r"check_ven_stop_trading.sql"
file_sql_cd_check_prgx = r"cd_check_prgx.sql"
file_sql_cd_ref_listagg = r"cd_ref_listagg.sql"
file_sql_cd_ref_listagg_item = r"cd_ref_listagg_item.sql"
file_sql_check_prof = r"check_profectus_detail.sql"

path_check_list = fr"D:\\python\\claim_pack_wrong_rate\\wrong_rate\\{folder_name}\\checklist.csv"
path_check_list_promo = fr"D:\\python\\claim_pack_wrong_rate\\wrong_rate\\{folder_name}\\check_list_promo.xlsx"

path_export = fr"D:\\python\\claim_pack_wrong_rate\\wrong_rate\\{folder_name}\\"
path_excel = r"CS_SCAN_Vendorname_Analyst_Date.xlsx"
path_dna = r"DNA.xlsx"
path_vba = r'CS_SCAN_vendorname_analyst_yyyymmdd.xlsb'


iconPath_email = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
iconPath_excel = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"

try:
    os.remove(r'D:\python\claim_pack_wrong_rate\claim_pack_wrong_rate.log')
except:
    pass

logging.basicConfig(filename="claim_pack_scan.log",
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    filemode='a')
 
# Creating an object
logger = logging.getLogger()
 
# Setting the threshold of logger to DEBUG
logger.setLevel(logging.INFO)


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
def connect_sql(cursor,file_sql,item_code=0,var_1=0,var_2=0,var_3=0,var_4=0,var_5=''):
    try:
        # cursor.execute((open(file_sql).read()))
        cursor.execute((open(file_sql).read()).format(item_code,var_1,var_2,var_3,var_4,var_5))
        all_rows = cursor.fetchall()
        field_names = [i[0] for i in cursor.description]
    finally:
        pass
        # conn.close()
    df = pd.DataFrame(all_rows)
    try:
        df.columns = field_names
    except ValueError:
        return pd.DataFrame(columns= field_names)
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

def get_info(df_splited):
    supp_num_list = list(df_splited['SUPPLIER'].drop_duplicates())
    item_list = list(df_splited['SKU_ID'].drop_duplicates())

    supp_num_list_final = convert_to_input_sql(num_list = supp_num_list)
    item_list_final = convert_to_input_sql(num_list = item_list)
    item_input_function = convert_to_input_function(num_list = item_list)
    return supp_num_list_final,item_list_final,item_input_function

def writer_excel(data,remove,number_sheet,path_export_final):
    # data = list_data, remove = list_remove,number_sheet= str(index_promo)+'_'+str(gst),path_export_final=path_export_final
    #select sheet
    print('Start write excel')
    sheet_df_mapping = {number_sheet: data}
    sheet_df_remove  = {number_sheet: remove}
    # Open Excel in background
    with xw.App(visible=False) as app:
        wb = app.books.open(path_export_final)
        # List of current worksheet names
        current_sheets = [sheet.name for sheet in wb.sheets]
        # Iterate over sheet/df mapping
        # If sheet already exist, overwrite current cotent. Else, add new sheet
        for sheet_name in sheet_df_mapping.keys():
            if sheet_name in current_sheets:
                for df_data in data :
                    wb.sheets(sheet_name).range(df_data['cell_export']).options(index=False,header=False).value = df_data['df']
            else:
                'Name of sheet cannot be found in Excel file, please check again'
        for sheet_name in sheet_df_remove.keys():
            if sheet_name in current_sheets:
                for df_remove in remove :
                    # wb.sheets(sheet_name).range(df_cell['cell_export']).options(index=False,header=False).value = df_cell['df']
                    length_start = df_remove['length_start'] + df_remove['count_df']
                    range_length_to_remove = str(length_start)+':'+ str(df_remove['length_end'])
                    wb.sheets(sheet_name).range(range_length_to_remove).api.Delete(DeleteShiftDirection.xlShiftUp)
            else:
                'Name of sheet cannot be found in Excel file, please check again'
        wb.save(path_export_final)
    print('Done write excel')
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

def create_worksheet(index_promo,gst,path_export_final,classify_state):
    # Open Excel in background
    with xw.App(visible=False) as app:
        if index_promo == 1:
            wb_from = app.books.open(path_excel)
        else :
            wb_from = app.books.open(path_export_final)
        if 'wrong' in classify_state.lower():
            ws_from = wb_from.sheets['template_wrong']
        else:
            ws_from = wb_from.sheets['template_entirely']
        ws_temp = wb_from.sheets['template_wrong']
        ws_from.copy(before=ws_temp, name=str(index_promo)+'_'+str(gst))
        wb_from.save(path_export_final)
    return 'Done create worksheet'     

def remove_sheet_change_xlsb(sheet_name_list,path_export_final,path_export_final_xlsb):
    print('Start delete sheet & change to xlsb')
    with xw.App(visible=False) as app:
        wb = app.books.open(path_export_final)   
        for sheet_name in sheet_name_list:        
            wb.sheets[sheet_name].delete()
        wb.save(path_export_final_xlsb)
    try:
        os.remove(path_export_final)
    except Exception as e:
        print(e)
    return print('Done delete sheet & change to xlsb')

# item_code=0,var_1=0,var_2=0,var_3=0,var_4=0

def df_sales_data(cursor, item_list_dict_gsted,start_date,end_date,classify_state):
    i = 0
    print(classify_state)
    # if classify_state == 'TO QA_STATE_SP' or  classify_state == 'TO QA_STATE_TH' :
    if 'STATE' in classify_state :
        for key,value in item_list_dict_gsted.items():
            print(key,value)
            item_code,state = key
            # if classify_state == 'TO QA_STATE_SP':
            if '_SP' in classify_state :
                df_each_item = connect_sql(cursor,file_sql = file_sql_summ_state ,item_code = item_code,var_1 = start_date,var_2 =end_date,var_3=value[0],var_4=value[1],var_5 = state)
            else:
                df_each_item = connect_sql(cursor,file_sql = file_sql_summ_state_th ,item_code = item_code,var_1 = start_date,var_2 =end_date,var_3=value[0],var_4=value[1],var_5 = state)
            if i == 0:
                df_merge = df_each_item
            else :
                df_merge = pd.concat([df_merge, df_each_item], ignore_index=True)
            i+=1
        # df_merge['ELI_CLAIM'] = df_merge.RQTY_PROMO * df_merge.SCAN_RATE
        df_merge= df_merge.sort_values(by=['RSKU_ID','RDAY_DT','RSTATE'], ascending=True).reset_index(drop=True)
    else:
        for key,value in item_list_dict_gsted.items():
            print(key,value)
            # if classify_state == 'TO QA_NATIONAL_SP':
            if '_SP' in classify_state:
                df_each_item = connect_sql(cursor,file_sql = file_sql_summ ,item_code = key,var_1 = start_date,var_2 =end_date,var_3=value[0],var_4=value[1])
            else:
                df_each_item = connect_sql(cursor,file_sql = file_sql_summ_th ,item_code = key,var_1 = start_date,var_2 =end_date,var_3=value[0],var_4=value[1])
            if i == 0:
                df_merge = df_each_item
            else :
                df_merge = pd.concat([df_merge, df_each_item], ignore_index=True)
            i+=1
        # df_merge['ELI_CLAIM'] = df_merge.RQTY_PROMO * df_merge.SCAN_RATE
        df_merge= df_merge.sort_values(by=['RSKU_ID','RDAY_DT','RSTATE'], ascending=True).reset_index(drop=True)
    return df_merge

def product_state_summary(df,df_ref):
    print('Start product_state_summary')
    list_data = []
    list_remove = []
    # Find distict item_code and state
    # writer_excel(df,cell_export,length_start,count_df,length_end,number_sheet,path_export_final)
    df_temp =df.drop_duplicates(['RSKU_ID','RITEM_DESC','RSTATE'])[['RSKU_ID','RITEM_DESC','RSTATE']]
    df_temp_2 = pd.merge(df_temp,df_ref,left_on=['RSKU_ID','RSTATE'],right_on=['ITEM_IDNT','CLM_STATE'], how='left')
    df_final = df_temp_2[['RSKU_ID','RITEM_DESC','RSTATE','REF_NUM','CLM_QTY','CLM_RATE','CLM_PRODUCT']]
    df_sku_desc = df_final[['RSKU_ID','RITEM_DESC']]
    df_state = df_final[['RSTATE']]
    df_ref = df_final[['REF_NUM','CLM_QTY','CLM_RATE','CLM_PRODUCT']]
    df_ref.insert(1,"REF_DESC",'')
    # Calculate number of rows
    number_rows_state = len(df_ref)

    dict_data_sku = {'df':df_sku_desc,'cell_export':'B217'}
    dict_data_state = {'df':df_state,'cell_export':'E217'}
    dict_data_remove = {'df':df_ref,'cell_export':'M217'}
    dict_remove = {'count_df':number_rows_state,'length_start':217,'length_end':1657}
    list_data.append(dict_data_sku)
    list_data.append(dict_data_state)
    list_data.append(dict_data_remove)
    list_remove.append(dict_remove)
    print('Done product_state_summary')
    return list_data,list_remove

def product_summary(df,df_item_ref):
    print('Start product_summary')
    list_data = []
    list_remove = []
    df_product =df.drop_duplicates(['RSKU_ID','RITEM_DESC'])[['RSKU_ID','RITEM_DESC']]
    df_temp = pd.merge(df_product,df_item_ref,left_on=['RSKU_ID'],right_on=['ITEM_IDNT'], how='left')
    # df_final = df_temp[['RSKU_ID','RITEM_DESC','REF_NUM']]
    df_product_1 = df_temp[['RSKU_ID','RITEM_DESC']]
    df_ref_1 = df_temp[['REF_NUM','CLM_QTY','CLM_RATE','CLM_PRODUCT']]
    df_ref_1.insert(1,"REF_DESC",'')
    number_rows_sales = len(df_product)
    # writer_excel(df = df_product,path_export_final = path_export_final, cell_export = 'B20',number_sheet = number_sheet,length_start=20 , count_df=number_rows_sales, length_end=116)
    dict_data_sku = {'df':df_product_1,'cell_export':'B20'}
    dict_data_ref = {'df':df_ref_1,'cell_export':'L20'}
    dict_remove = {'count_df':number_rows_sales,'length_start':20,'length_end':212}
    list_data.append(dict_data_sku)
    list_data.append(dict_data_ref)
    list_remove.append(dict_remove)
    print('Done product_summary')
    return list_data , list_remove

def cd_ref(prmt_id,cursor,file_sql,item_code,df_sales, file_sql_2,file_sql_3,classify_state):
    print('Start cd ref')
    list_data = []
    list_remove = []
    # df_sales['ELI_CLAIM'] = df_sales.RQTY_PROMO * df_sales.SCAN_RATE
    df_sales_export = df_sales
    # print(df_sales_export)
    df_ref = connect_sql(cursor=cursor,file_sql = file_sql ,item_code = item_code,var_1 = prmt_id)
    df_ref_groupby = df_ref.groupby('CLM_REF_NUM').agg({'CLM_PRODUCT':'sum'}).sort_values(by='CLM_PRODUCT', ascending=True).reset_index()
    df_sales_daily = pd.concat([df_sales, df_ref_groupby], axis=1 )
    print('Done cd ref')
    print('start state ref')
    df_state_ref = connect_sql(cursor=cursor,file_sql = file_sql_2 ,item_code = item_code,var_1 = prmt_id)
    print('done state ref')
    print('start item ref')
    df_item_ref = connect_sql(cursor=cursor,file_sql = file_sql_3 ,item_code = item_code,var_1 = prmt_id)
    print('done item ref')
    # writer_excel(df = df_sales, cell_export = 'B174',number_sheet= str(index_promo)+'_'+str(gst),length_start=174 ,count_df=len(df_sales), length_end=10174,path_export_final=path_export_final)
    if 'entirely' in classify_state.lower():
        df_sales_export_1 = df_sales_export[['RDAY_DT', 'RSKU_ID', 'RITEM_DESC', 'RSTATE', 'RF_NORM_SELL_PRICE_AMT','RPRM_PRICE', 'RTOTAL_SALES_AMT_MOD', 'RQTY','RF_ACTUAL_SELL_PRICE_AMT_MOD', 'RTOTAL_SALES_AMT_NON', 'RQTY_NON','RF_ACTUAL_SELL_PRICE_AMT_MOD_NON', 'RTOTAL_SALES_AMT_MOD_PROMO','RQTY_PROMO', 'RF_ACTUAL_SELL_PRICE_AMT_MOD_1', 'SCAN_RATE']]
        dict_data_sales = {'df':df_sales_export_1,'cell_export':'B1662'}
        list_data.append(dict_data_sales)
        df_sales['ELI_CLAIM'] = df_sales.RQTY_PROMO * df_sales.SCAN_RATE      
    else:
        df_sales_export_1 = df_sales_export[['RDAY_DT', 'RSKU_ID', 'RITEM_DESC','RSTATE', 'RF_NORM_SELL_PRICE_AMT', 'RPRM_PRICE']]
        df_sales_export_2 = df_sales_export[['RQTY']]
        df_sales_export_3 = df_sales_export[['SCAN_RATE']]
        dict_data_sales_1 = {'df':df_sales_export_1,'cell_export':'B1662'}
        dict_data_sales_2 = {'df':df_sales_export_2,'cell_export':'I1662'}
        dict_data_sales_3 = {'df':df_sales_export_3,'cell_export':'Q1662'}
        list_data.append(dict_data_sales_1)
        list_data.append(dict_data_sales_2)
        list_data.append(dict_data_sales_3)
        df_sales['ELI_CLAIM'] = df_sales.RQTY * df_sales.SCAN_RATE
    dict_data_sales_ref = {'df':df_ref_groupby,'cell_export':'S1662'}
    dict_remove = {'count_df':len(df_sales),'length_start':1662,'length_end':121662}
    # list_data.append(dict_data_sales)
    list_data.append(dict_data_sales_ref)
    list_remove.append(dict_remove)
    # df_sales.to_csv(r'D:\OneDrive - Profectus Group\Desktop\test.csv')
    sum_sales = df_sales['ELI_CLAIM'].sum() 
    if not df_ref_groupby.empty:
        sum_ref = df_ref_groupby['CLM_PRODUCT'].sum()
    else:
        sum_ref = 0
    gap_sales_ref = sum_sales - sum_ref
    print(sum_sales,sum_ref,gap_sales_ref)
    return df_item_ref,df_state_ref,df_ref,df_sales_daily,list_data,list_remove,gap_sales_ref

def insert_attachments(sheet_name,file_path_excel,file_path_email,path_export_final):  
    print('Start insert email and excel')
    print(file_path_excel)
    print(file_path_email)
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    xl.Visible = False
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(path_export_final, UpdateLinks = True)
    ws = wb.Worksheets(sheet_name)
    try:
        excel_name = file_path_excel.split('/')[-1][0:10]
    except:
        excel_name ='excel'
    try:
        email_name = file_path_email.split('/')[-1][0:10]
    except:
        email_name = 'email'
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

def main():
    cursor = set_up(config = config_coles)
    logging.info('-----------------------------------START CHECK COLUMN-----------------------------------------------------')
    df_raw_check = connect_sql(cursor = cursor,file_sql = file_sql_claimpack,item_code=month_filter)
    logging.info('----------------------------------DONE CHECK COLUMN------------------------------------------------------------------')
    #Export checklist
    # df_raw_check.drop(columns = ['FILTER_ITEM']).to_csv(path_check_list,index=False)
    # df_raw_check.to_csv(path_check_list,index=False)
    if not os.path.isfile(os.path.join(current_dir,'wrong_rate',folder_name,path_check_list_xlsx)):
        df_raw_check.to_csv(path_check_list,index=False)
        print('No checklist excel exist')
        return None
    

    # Check with analyzed excel file
    path_check_list_excel_update = path_check_list_xlsx.split('.')[0] + '_updated.csv'
    df_csv = df_raw_check
    df_xlsx = pd.read_excel(os.path.join(current_dir,'wrong_rate',folder_name,path_check_list_xlsx))
    df_xlsx.applymap(lambda x : str(x))
    df_xlsx['TO_EXPORT'] = df_xlsx['TO_EXPORT'].astype(str).replace('nan','')
    df_xlsx_distinct = df_xlsx[['PRMTN_COMP_IDNT','SUPPLIER','DEPT_IDNT','CHECK_COLUMN','TO_EXPORT','NOTE']].drop_duplicates().reset_index(drop=True).applymap(lambda x : str(x))
    df_xlsx_distinct_to_map = df_xlsx_distinct[df_xlsx_distinct['TO_EXPORT'].str.contains('YES') | df_xlsx_distinct['TO_EXPORT'].str.contains('NO')].reset_index(drop=True)
    df_merged = pd.merge(how='left',left = df_csv, right = df_xlsx_distinct_to_map , on= ['PRMTN_COMP_IDNT','SUPPLIER','DEPT_IDNT','CHECK_COLUMN'])
    df_merged['TO_EXPORT'] = df_merged['TO_EXPORT'].astype(str).replace('nan','')
    df_merged['NOTE'] = df_merged['NOTE'].astype(str).replace('nan','')
    df_merged.to_csv(os.path.join(current_dir,'wrong_rate',folder_name,path_check_list_excel_update),index = False)
    # return 0

    ###############################################################
    error_list =[]
    time_start = datetime.datetime.now()
    # Filter df_splited with condition , keep TO QA and PRGX
    # df_raw_filter = df_raw_check[(df_raw_check['CHECK_COLUMN'] == 'TO QA_STATE')] 
    # df_raw_filter = df_merged[df_merged['CHECK_COLUMN'].str.contains('TO QA') & df_merged['TO_EXPORT'] == 'YES']
    df_raw_filter = df_merged[(df_merged['TO_EXPORT'] == 'YES') & df_merged['CHECK_COLUMN'].str.contains('TO QA')]
    # df_raw_filter = df_raw_check[(df_raw_check['CHECK_COLUMN'] == 'TO QA_NATIONAL')| (df_raw_check['CHECK_COLUMN'] == 'TO QA_STATE')] 
    df_unique_supp_filter = df_raw_filter[['SUPPLIER','PRMTN_COMP_IDNT','CHECK_COLUMN','DEPT_IDNT']].drop_duplicates().values.tolist()
    # Create dictionary with supp_num key and list of promo_ids after filter conditions, keep check again and to QA
    dict_sup_pro_filter = {}
    j=0
    for list_sup in df_unique_supp_filter:
        if j == 0:
            dict_sup_pro_filter[list_sup[0]] = [[list_sup[1]] + [list_sup[2]]+ [list_sup[3]]]
        else:
            if list_sup[0] in dict_sup_pro_filter.keys():
                dict_sup_pro_filter[list_sup[0]].append([list_sup[1]] + [list_sup[2]]+ [list_sup[3]])
            else:
                dict_sup_pro_filter[list_sup[0]] = [[list_sup[1]] + [list_sup[2]]+ [list_sup[3]]]
        j+=1
    df_raw_filter['UNIT_PRICE_TTL'] = df_raw_filter['UNIT_PRICE_TTL'].astype('float').round(2)
    df_raw_filter['ELI_MAX_RATE_CD_SUM'] = df_raw_filter['ELI_MAX_RATE_CD_SUM'].astype('float')
    df_raw_filter['ELI_MOD_RATE_CD_SUM'] = df_raw_filter['ELI_MOD_RATE_CD_SUM'].astype('float')
    dict_splitted = {}
    for key,value in dict_sup_pro_filter.items():
        if(len(value) <= 5):
            dict_splitted[key] = value
        else:
            max_index_supplier  = math.ceil(len(value) /5)
            for i in range(1,max_index_supplier+1):
                dict_splitted[f'{key}_v{i}'] = value[5*(i-1):5*i]
    print(dict_splitted)
    check_list_promo_index = 1
    for supp_num,list_pmt_id_classify in dict_splitted.items():
        index_promo=1
        # summary_index = 1
        summary_index_list = []
        logging.info(f'-------------------------------------------------Working on supp_num {supp_num}-----------------------------------------------------')
        print(f'-------------------------------------------------Working on supp_num {supp_num} -----------------------------------------------------')
        for pmt_id_classif in list_pmt_id_classify:
            classify_state = pmt_id_classif[1]
            pmt_id = pmt_id_classif[0]
            dept = pmt_id_classif[2]
            print(f'---------------------------------------------Working on prmtn_id {pmt_id} ---------------------------------------------------------')
            check_list_promo = []
            supp_num_splitted = supp_num.split('_')[0]
            print(supp_num_splitted)
            df_splited_filter = df_raw_filter[(df_raw_filter['SUPPLIER'] == supp_num_splitted) & (df_raw_filter['PRMTN_COMP_IDNT'] == pmt_id) & (df_raw_filter['FILTER_ITEM'] == 'YES') & (df_raw_filter['DEPT_IDNT'] == dept) ] 
            # get some important variable
            supp_num_list_final,item_list_final,item_input_function = get_info(df_splited = df_splited_filter)
            print(df_splited_filter)
            supp_desc = df_splited_filter['SUPP_DESC'].unique()[0].replace("/","")
            sc = ['.', '>', '<', '!', '?', ',', '[', ']', '(', ')', '&', "'", '"', '+', ':', ';']
            supp_desc = ''.join([c for c in supp_desc if c not in sc])
            supp_desc = supp_desc.strip()
            print('supp_desc',supp_desc)
            gst = int(df_splited_filter['CML_COST_GST_RATE_PCT'].unique()[0])
            dept_desc = df_splited_filter['DEPT_DESC'].unique()[0]
            prmt_id = df_splited_filter['PRMTN_COMP_IDNT'].unique()[0]
            prmt_name = df_splited_filter['PRMTN_COMP_NAME'].unique()[0] 
            logging.info(f'----------------------------------------Working on supp_desc {supp_desc}, prmtn id {pmt_id}------------------------------------------')
            try:
                paf_loc = df_splited_filter['PAF_LOCATION'].unique()[0]
            except Exception :
                paf_loc = '0'
            try:
                email_loc = df_splited_filter['EMAIL'].unique()[0]
            except Exception :
                email_loc = '0'
            vendor_num = df_splited_filter['VENDOR_NUM'].unique()[0] 
            print(df_splited_filter['PRMTN_COMP_DTL_SDT'].unique()[0])
            clm_start = df_splited_filter['PRMTN_COMP_DTL_SDT'].unique()[0]
            clm_end = df_splited_filter['PRMTN_COMP_DTL_END_DT'].unique()[0]
            clm_start_converted = clm_start.strftime('%d/%m/%Y')
            clm_end_converted = clm_end.strftime('%d/%m/%Y')
            clm_start_converted_promo = clm_start.strftime('%Y-%m-%d')
            clm_end_converted_promo = clm_end.strftime('%Y-%m-%d')
            amt_max = df_splited_filter['ELI_MAX_RATE_CD_SUM'].unique()[0].astype(str)
            amt_mod = df_splited_filter['ELI_MOD_RATE_CD_SUM'].unique()[0].astype(str)
            max_process_date = df_splited_filter['MAX_PROCESS_DTE'].unique()[0]
            min_start_date = str(df_splited_filter['MAX_START_DATE'].min())
            max_end_date = str(df_splited_filter['MAX_END_DATE'].max())
            #create path for excel and path_xlsb for excel
            path_export_final = path_export+'CS_SCAN_'+supp_desc+'_Analyst_date.xlsx'
            path_export_final_xlsb = path_export+'CS_SCAN_'+supp_desc+'_Analyst_date_'+supp_num+'.xlsb'
            # sc = ['.', '>', '<', '!', '?', ',', '[', ']', '(', ')', '&', "'", '"', '+', ':', ';']
            # path_export_final = ''.join([c for c in path_export_final if c not in sc])
            # path_export_final = path_export_final.strip()
            # path_export_final_xlsb = ''.join([c for c in path_export_final_xlsb if c not in sc])
            # path_export_final_xlsb = path_export_final_xlsb.strip()
            # print('path_export_final',path_export_final)
            create_worksheet(index_promo=index_promo,gst=gst,path_export_final=path_export_final,classify_state=classify_state)
            # if classify_state == 'TO QA_NATIONAL_SP' or classify_state == 'TO QA_NATIONAL_TH' :
            if 'NATIONAL' in classify_state:
                # df_splited_filter = df_splited_filter[['SKU_ID','STATE','UNIT_PRICE_TTL','MAX_RATE']].drop_duplicates()
                df_splited_filter = df_splited_filter[['SKU_ID','UNIT_PRICE_TTL','MAX_RATE']].drop_duplicates()
                df_splited_filter = df_splited_filter.groupby(by = ['UNIT_PRICE_TTL','MAX_RATE'])['SKU_ID'].agg(list).to_frame().reset_index()
                df_splited_filter['SKU_ID'] = df_splited_filter['SKU_ID'].apply(lambda x : convert_to_input_function(x))
                item_list_dict = df_splited_filter.set_index('SKU_ID')[['UNIT_PRICE_TTL','MAX_RATE']].to_dict('index')
                for key,value in item_list_dict.items():
                    item_list_dict[key] = [item_list_dict[key]['UNIT_PRICE_TTL']] + [item_list_dict[key]['MAX_RATE']] 
                print(item_list_dict)
            else:
                df_splited_filter = df_splited_filter[['SKU_ID','STATE','UNIT_PRICE_TTL','MAX_RATE']].drop_duplicates()
                df_splited_filter = df_splited_filter.groupby(by = ['UNIT_PRICE_TTL','MAX_RATE','STATE'])['SKU_ID'].agg(list).to_frame().reset_index()
                df_splited_filter['SKU_ID'] = df_splited_filter['SKU_ID'].apply(lambda x : convert_to_input_function(x))
                df_splited_filter = df_splited_filter.groupby(by = ['UNIT_PRICE_TTL','MAX_RATE','SKU_ID'])['STATE'].agg(list).to_frame().reset_index()
                df_splited_filter['STATE'] = df_splited_filter['STATE'].apply(lambda x : convert_to_input_sql(x))
                item_list_dict = df_splited_filter.set_index(['SKU_ID','STATE'])[['UNIT_PRICE_TTL','MAX_RATE']].to_dict('index')
                for key,value in item_list_dict.items():
                    item_list_dict[key] = [item_list_dict[key]['UNIT_PRICE_TTL']] + [item_list_dict[key]['MAX_RATE']] 
                    print(item_list_dict) 
            # To create excel file
            summary_index_list.append(str(index_promo)+'_'+str(gst))
            # writer_excel_without_remove_rows(df = df_splited,path = path_export_final,cell_export = 'A1',number_sheet=str(index_promo)+'_'+str(gst),path_export_final=path_export_final)   
            df_sales = df_sales_data(cursor = cursor,item_list_dict_gsted = item_list_dict ,start_date = clm_start_converted,end_date = clm_end_converted,classify_state=classify_state) 
            print(f"------------------------------------------len(df_sales) : {df_sales['RSTATE'].count()}---------------")  
            # df_ref,df_sales = cd_ref(prmt_id=prmt_id,cursor = cursor,file_sql=file_sql_cd_ref,item_code=item_list_final,df_sales=df_sales)  
            df_item_ref,df_state_ref,df_ref,df_sales,list_data_sales,list_remove_sales,gap_sales_ref  = cd_ref(prmt_id=prmt_id,cursor = cursor,file_sql=file_sql_cd_ref,item_code=item_list_final,df_sales=df_sales,file_sql_2 = file_sql_cd_ref_listagg, file_sql_3 = file_sql_cd_ref_listagg_item,classify_state=classify_state)  
            # return 0
            try:
                df_ref_groupby = df_ref.groupby('CLM_REF_NUM').agg({'CLM_PRODUCT':'sum'}).sort_values(by='CLM_PRODUCT', ascending=True).reset_index()
                ref_num_list = ', '.join(df_ref_groupby['CLM_REF_NUM'].tolist())
            except:
                ref_num_list = ''
            list_data_state,list_remove_state = product_state_summary(df = df_sales,df_ref=df_state_ref)
            list_data_product ,list_remove_product  = product_summary(df = df_sales, df_item_ref = df_item_ref)
            dict_data_dept = {'df':dept,'cell_export':'F8'}
            dict_data_prmt_id = {'df':prmt_id,'cell_export':'B12'}
            dict_data_prmt_name = {'df':prmt_name,'cell_export':'C12'}
            # dict_data_paf_loc = {'df':paf_loc,'cell_export':'J4'}
            # dict_data_email_loc = {'df':email_loc,'cell_export':'K4'}
            dict_data_supp_num = {'df':supp_num_splitted,'cell_export':'E8'}
            dict_data_supp_desc = {'df':supp_desc,'cell_export':'C8'}
            dict_data_vendor_num = {'df':vendor_num,'cell_export':'D8'}
            dict_data_claim_number = {'df':str(index_promo)+'_'+str(gst),'cell_export':'B16'}
            dict_data_ref_num_list = {'df':ref_num_list,'cell_export':'D16'}
            list_data = list_data_sales + list_data_state + list_data_product + [dict_data_dept] + [dict_data_prmt_id] + [dict_data_prmt_name]  + [dict_data_supp_num] + [dict_data_supp_desc] + [dict_data_vendor_num] + [dict_data_claim_number] + [dict_data_ref_num_list]
            list_remove = list_remove_sales + list_remove_state + list_remove_product
            #  Fill sheet Complete Daily Sales Data
            writer_excel(data = list_data, remove = list_remove,number_sheet= str(index_promo)+'_'+str(gst),path_export_final=path_export_final) 
            # try: 
            #     insert_attachments(sheet_name = str(index_promo)+'_'+str(gst),file_path_excel = paf_loc ,file_path_email = email_loc,path_export_final = path_export_final)
            #     # logging.warning('---------------------------Cannot insert attachments------------------------------------------------------------')
            # except:
            #     logging.warning('---------------------------Cannot insert attachments------------------------------------------------------------')
            # insert_attachments(sheet_name = str(index_promo)+'_'+str(gst),file_path_excel = paf_loc ,file_path_email = email_loc,path_export_final = path_export_final)
            #export check_list_promo
            index_promo+=1
            check_list_promo_index += 1 
            with xw.App(visible=False) as app:
                print('check_list_promo_index')
                if check_list_promo_index == 2:
                    wb = app.books.open('check_list_promo.xlsx')
                else:
                    wb = app.books.open(path_check_list_promo)
                wb_sheet = wb.sheets['Sheet1']
                check_list_promo = [supp_num] + [supp_desc] +[str(index_promo-1)+'_'+str(gst)] +[prmt_id] + [clm_start_converted_promo] + [clm_end_converted_promo]+ [dept] +[dept_desc] + [amt_max] + [amt_mod] + [max_process_date] + [min_start_date] + [max_end_date] + [gap_sales_ref] + [classify_state] +['Done'] + [email_loc] + [paf_loc]
                print(check_list_promo)
                wb_sheet.range(f'A{check_list_promo_index}').value =  check_list_promo
                wb.save(path_check_list_promo) 
            # index_promo+=1
            # check_list_promo_index += 1    
            # return 0          
        # Fill sheet Vendor Summary
        fill_summary_sheet(summary_index_list= summary_index_list,path_export_final=path_export_final)         
        # remove_sheet_change_xlsb(sheet_name_list = ['template_wrong','template_entirely'],path_export_final=path_export_final ,path_export_final_xlsb = path_export_final_xlsb)
        # move_worksheet_to_vba_template(path_xlsb=path_export_final_xlsb)  
        print('-------------------------------------------------------------------------------------------------------------------------------------')
    return None

if __name__ == '__main__':
    if not os.path.isdir(fr"D:\\python\\claim_pack_wrong_rate\\wrong_rate\\{folder_name}"):
        os.mkdir(fr"D:\\python\\claim_pack_wrong_rate\\wrong_rate\\{folder_name}")
    main()