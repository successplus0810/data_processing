import json
import pandas as pd
import snowflake.connector as sf
import os
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
import datetime
import math
import win32com.client as win32
from pywintypes import com_error
pd.options.mode.chained_assignment = None
import logging
import numpy as np

##########################
analyst_name = 'Analyst'
date_batch = 'Date'
#########################
iconPath_email = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
iconPath_excel = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"

######
os.chdir(r'D:\python\claim_pack_cl')
try:
    os.remove(r'D:\\python\cl_summarizer\claim_pack_cl.log')
except:
    pass

logging.basicConfig(filename="claim_detail.log",
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    filemode='a')
 
# Creating an object
logger = logging.getLogger()
 
# Setting the threshold of logger to DEBUG
logger.setLevel(logging.INFO)
###########
number_sheet_ap = 'AP'
number_sheet_pd = 'Promo Detail'
number_sheet_cd = 'Claims Detail'

config_coles = r"config.json"

file_sql_claim_pack = r"claim_pack.sql"
file_sql_pd = r"pd.sql"
file_sql_cd = r"cd.sql"
file_sql_ap = r"ap.sql"
file_sql_dept = r"dept.sql"
file_sql_gst = r"gst.sql"
file_sql_summ = r"summarizer.sql"
# file_sql_cd_ref = r"cd_ref.sql"
file_sql_summarizer_state_single = r"summarizer_state_single.sql"
file_sql_summarizer_state_bundle = r"summarizer_state_bundle.sql"
file_sql_summarizer_national_single = r"summarizer_national_single.sql"
file_sql_summarizer_national_bundle = r"summarizer_national_bundle.sql"
file_sql_cd_national = r"cd_national.sql"
file_sql_cd_state = r"cd_state.sql"
file_sql_check_category_name = r"category_name.sql"
file_sql_check_category_id= r"category_id.sql"
file_sql_get_ven_id_name = r"get_ven_id_name.sql"
file_sql_check_prof = r"check_prof.sql"

folder_name = '202111'
###############################################
config_coles = r"config.json"

path_excel = 'CL_SCAN_Vendorname_Analyst_Date.xlsx'
# path_import_item = 'item_import_1.xlsx'
path_vba = 'CL_SCAN_vendorname_analyst_yyyymmdd_LESSTHAN20K.xlsb'

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
cursor = set_up(config = config_coles)

class claim_pack_cl:
    def __init__(self,cursor,month_filter):
        self.cursor = cursor
        self.month_filter = month_filter

    def connect_sql(self,file_sql,var_1='',var_2='',var_3='',var_4='',var_5='',var_6='',var_7='',var_8='',var_9=''):
        try:
            self.cursor.execute((open(file_sql).read()).format(var_1,var_2,var_3,var_4,var_5,var_6,var_7,var_8,var_9))
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
    
    def check_profectus_claim(self,df):
        list_excel = df.to_dict('records')
        list_prof = []
        for row in list_excel:
            df_check_prof = self.connect_sql(file_sql=file_sql_check_prof,var_1 = row['DEAL'] ,var_2 = row['ITEMIDSKU'],var_3 = row['STARTDATE'],var_4= row['ENDDATE'],var_5 = row['BRANDID'], var_6 = row['UOM'],var_7 = row['STARTDATE'], var_8 =row['ENDDATE'],var_9 = row['STATE'])
            if not df_check_prof.empty: 
                row['ITEM_RAISED'] = df_check_prof['ITEM_RAISED'][0]
                row['CLAIM_PROF'] = df_check_prof['CLAIM_PROF'][0]
                row['FILE_PATH'] = df_check_prof['FILE_PATH'][0]
                list_prof.append(row)
            else:
                list_prof.append(row)
        df_check_prof_detail = pd.DataFrame.from_records(list_prof).reset_index(drop=True)
        return df_check_prof_detail
    
    def create_check_column_for_checklist(self):
        df_raw = self.connect_sql(file_sql=file_sql_claim_pack , var_1= self.month_filter)
        df_raw['KEY'] = df_raw[['REBATEDATE','BRANDID','UOM','STARTDATE','ENDDATE', 'CLASSIFY_STATE','CLASSIFY_PROMO']].astype('str').apply(lambda row: '+'.join(row.values), axis=1)
        df_unique_supp = df_raw[['VENDOR_NUM','KEY']].drop_duplicates().values.tolist()
        dict_ven_promo = {}
        i=0
        for list_sup in df_unique_supp:
            if i == 0:
                dict_ven_promo[list_sup[0]] = [list_sup[1]]
            else:
                if list_sup[0] in dict_ven_promo.keys():
                    dict_ven_promo[list_sup[0]].append(list_sup[1])
                else:
                    dict_ven_promo[list_sup[0]] = [list_sup[1]]
            i+=1
        # Initialize df_raw_check as an empty DataFrame
        if dict_ven_promo == {}:
            df_raw_check = pd.DataFrame(columns= df_raw.columns)
        else:
            j = 0
            for vendor_num,list_key in dict_ven_promo.items():
                # To classify Check_column
                for key in list_key:
                    df_splited = df_raw[(df_raw['VENDOR_NUM'] == vendor_num) & (df_raw['KEY'] == key)] 
                    if df_splited['ELI'].iloc[0] < 100:
                        df_splited['CHECK_COLUMN'] = 'ELI < 100'
                    else:
                        df_check_prof_detail = self.check_profectus_claim(df = df_splited)
                        if 'ITEM_RAISED' in df_check_prof_detail.columns:
                            df_check_prof_detail['CHECK_COLUMN'] = 'PROFECTUS CLAIMED'
                            df_check_prof_detail.loc[(df_check_prof_detail['ITEM_RAISED'].isnull()), 'CHECK_COLUMN'] = 'PROFECTUS_CLAIMED_NOT_RAISED'
                            df_splited = df_check_prof_detail
                        else:
                            classify_state  = df_splited['CLASSIFY_STATE'].iloc[0]
                            classify_promo  = df_splited['CLASSIFY_PROMO'].iloc[0]
                            df_splited['CHECK_COLUMN'] = f'TO QA_{classify_state}_{classify_promo}'
                    print(df_splited)    
                    if j  == 0:
                        df_raw_check = df_splited
                    else:
                        df_raw_check = pd.concat([df_raw_check, df_splited], ignore_index=True)
                    j+=1
        return df_raw_check
    def export(self):
        month_filter_converted = self.month_filter.replace('-','_')
        df_export  = self.create_check_column_for_checklist()
        print(df_export)
        file_name = f'checklist_{month_filter_converted}'
        df_export.to_csv(fr'D:\python\claim_pack_cl\{file_name}.csv',index=False)
        # Load the data from your CSV file
        df_export = pd.read_csv(fr'D:\python\claim_pack_cl\{file_name}.csv')
        # Write the data to an Excel file
        df_export.to_excel(fr'D:\python\claim_pack_cl\{file_name}.xlsx', index=False)
        os.remove(fr'D:\python\claim_pack_cl\{file_name}.csv')
        return df_export

claim_pack_sep = claim_pack_cl(cursor,'2022-09')

# test.cursor
claim_pack_sep.export()