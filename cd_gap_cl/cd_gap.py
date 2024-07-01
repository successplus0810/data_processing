import json
import pandas as pd
import snowflake.connector as sf
import os
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
import datetime
import math
pd.options.mode.chained_assignment = None
###### Analyst fill
folder_name = '20220408'

###############################################
config_coles = r"config.json"

file_sql_claimpack = r"cd_gap.sql"
file_sql_summ = r"summarizer.sql"
# file_sql_cd_ref = r"cd_ref.sql"
file_sql_cd_check_again = r"cd_check_again.sql"
file_sql_ven_stop_trading = r"check_ven_stop_trading.sql"
# file_sql_cd_check_prgx = r"cd_check_prgx.sql"

path_check_list = fr"D:\\python\\cd_gap_cl\\national\\{folder_name}\\checklist.xlsx"
path_check_list_promo = fr"D:\\python\\cd_gap_cl\\national\\{folder_name}\\check_list_promo.xlsx"
path_check_list_test = fr"D:\\python\\cd_gap_cl\\national\\{folder_name}\\checklist_1.xlsx"

path_export = fr"D:\\python\\cd_gap_cl\\national\\{folder_name}\\"
path_excel = r"CL_SCAN_Vendorname_Analyst_Date.xlsx"
# path_dna = r"DNA.xlsx"

os.chdir("D:\\python\\cd_gap_cl")

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
def connect_sql(cursor,file_sql,item_code=0,var_1=0,var_2=0,var_3=0,var_4=0,var_5=0,var_6 =0):
    try:
        # cursor.execute((open(file_sql).read()))
        cursor.execute(open(file_sql).read().format(item_code,var_1,var_2,var_3,var_4,var_5 ,var_6))
        all_rows = cursor.fetchall()
        field_names = [i[0] for i in cursor.description]
    finally:
        pass
        # conn.close()
    df = pd.DataFrame(all_rows)
    try:
        df.columns = field_names
    except ValueError:
        return pd.DataFrame([])
    return df

def convert_to_input_sql(num_list):
    num_list_final = ''
    # print('SUPP LIST',supp_num_list)
    for num_list in num_list:
        num_list_final = num_list_final + "'" + str(num_list) + "',"
    return num_list_final[:-1]

def convert_to_input_function(num_list):
    num_list_final = ''
    # print('SUPP LIST',supp_num_list)
    for num_list in num_list:
        num_list_final = num_list_final + str(num_list) + ','
    return num_list_final[:-1]

def get_info(df_splited):
    supp_num_list = list(df_splited['VENDOR_NUMBER'].drop_duplicates())
    item_list = list(df_splited['ITEMID'].drop_duplicates())

    supp_num_list_final = convert_to_input_sql(num_list = supp_num_list)
    item_list_final = convert_to_input_sql(num_list = item_list)
    item_input_function = convert_to_input_function(num_list = item_list)
    return supp_num_list_final,item_list_final,item_input_function

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
    return None

def fill_summary_sheet(supp_desc,summary_index_list,path_export_final,vendor_num):
    print('Start fill summary sheet')
    with xw.App(visible=False) as app:
        wb_from = app.books.open(path_export_final)
        summary_index = 1
        for index in summary_index_list:
            wb_from.sheets['Supplier Summary'].range('B'+str(summary_index+7)).value = index
            wb_from.sheets['Supplier Summary'].range('D'+str(summary_index+7)).value = supp_desc
            summary_index += 1
            wb_from.sheets['Supplier Summary'].range('B'+str(summary_index+7)+':N'+str(summary_index+7)).clear_contents()
            wb_from.sheets['Supplier Summary'].range('B'+str(summary_index+7)+':N'+str(summary_index+7)).clear_formats()
            summary_index += 1
        length_start = summary_index + 7
        range_length_to_remove = str(length_start -1)+':'+ str(38)
        wb_from.sheets('Supplier Summary').range(range_length_to_remove).api.Delete(DeleteShiftDirection.xlShiftUp)  
        print('Done fill summary sheet')
        print('start AP sheet' )
        wb_from.sheets.add('AP')
        wb_from.sheets['AP'].range('A2').value = vendor_num 
        wb_from.save(path_export_final)
        print('Done AP sheet' )
    return None

def create_worksheet(index_promo,path_export_final):
    # Open Excel in background
    with xw.App(visible=False) as app:
        if index_promo == '1':
            wb_from = app.books.open(path_excel)
        else :
            wb_from = app.books.open(path_export_final)
        ws_from = wb_from.sheets['template']
        ws_from.copy(before=ws_from, name=index_promo)
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

# item_code=0,var_1=0,var_2=0,var_3=0,var_4=0

def df_sales_data(cursor , file_sql , item_list_dict_gsted,start_date,end_date):
    list_data = []
    list_remove = []
    i = 0
    for key,value in item_list_dict_gsted.items():
        item,brand,uomm = key
        promo_price, scan_rate = value
        uomm = int(uomm)
        df_each_item = connect_sql(cursor = cursor,file_sql = file_sql ,item_code = scan_rate,var_1 =item , var_2 = start_date,var_3=end_date,var_4=brand,var_5=uomm,var_6 =promo_price)
        if i == 0:
            df_merge = df_each_item
        else :
            df_merge = pd.concat([df_merge, df_each_item], ignore_index=True)
        i+=1
    df_merge['ELI_CLAIM'] = df_merge.ITEMQUANTITY_PROMO * df_merge.SCAN_RATE
    dict_data = {'df':df_merge,'cell_export':'B606'}
    dict_remove = {'count_df':len(df_merge),'length_start':606,'length_end':20606}
    list_data.append(dict_data)
    list_remove.append(dict_remove)
    return df_merge,list_data,list_remove

def product_state_summary(df_sales):
    print('Start product_state_summary')
    list_data = []
    list_remove = []
    # df: df_sales
    # Find distict item_code and state
    # writer_excel(df,cell_export,length_start,count_df,length_end,number_sheet,path_export_final)
    df_temp =df_sales.drop_duplicates(['ITEMIDSKU','ITEMNAME','BRANDID','UOM_QTY','STATE'])[['ITEMIDSKU','ITEMNAME','BRANDID','UOM_QTY','STATE']]
    df_temp.insert(2,"BLANK",'')
    # Calculate number of rows
    number_rows_state = len(df_temp)
    dict_data_sku = {'df':df_temp,'cell_export':'B111'}
    dict_remove = {'count_df':number_rows_state,'length_start':111,'length_end':601}
    list_data.append(dict_data_sku)
    list_remove.append(dict_remove)
    print('Done product_state_summary')
    return list_data,list_remove

def product_summary(df_sales):
    print('Start product_summary')
    list_data = []
    list_remove = []
    df_product =df_sales.drop_duplicates(['ITEMIDSKU','ITEMNAME','BRANDID','UOM_QTY'])[['ITEMIDSKU','ITEMNAME','BRANDID','UOM_QTY']]
    df_product.insert(2,"BLANK",'')
    number_rows_sales = len(df_product)
    # writer_excel(df = df_product,path_export_final = path_export_final, cell_export = 'B20',number_sheet = number_sheet,length_start=20 , count_df=number_rows_sales, length_end=116)
    dict_data = {'df':df_product,'cell_export':'B8'}
    dict_remove = {'count_df':number_rows_sales,'length_start':8,'length_end':104}
    list_data.append(dict_data)
    list_remove.append(dict_remove)
    print('Done product_summary')
    return list_data , list_remove

  
def main():
    cursor = set_up(config = config_coles)
    df_raw = connect_sql(cursor = cursor,file_sql = file_sql_claimpack)
    df_unique_supp = df_raw[['VENDOR_NUMBER','STARTDATE','ENDDATE','CLASSIFY_TYPE','CLASSIFY_CLAIM']].drop_duplicates().values.tolist()
    # Read vendor stop trading
    df_ven_stop_trading = connect_sql(cursor = cursor,file_sql = file_sql_ven_stop_trading)
    list_ven_stop_trading = df_ven_stop_trading['VENDOR_NUM'].drop_duplicates().values.tolist()
    # Create dictionary with supp_num key and list of promo_ids
    # print(df_unique_supp)
    # return 0
    dict_sup_pro = {}
    i=0
    for list_sup in df_unique_supp:
        if i == 0:
            dict_sup_pro[list_sup[0]] = [[list_sup[1]] + [list_sup[2]] +[list_sup[3]]+[list_sup[4]]]
        else:
            if list_sup[0] in dict_sup_pro.keys():
                dict_sup_pro[list_sup[0]].append([list_sup[1]] + [list_sup[2]] +[list_sup[3]]+[list_sup[4]])
                # dict_sup_pro[list_sup[0]].append(list_sup[2])
            else:
                dict_sup_pro[list_sup[0]] = [[list_sup[1]] + [list_sup[2]] +[list_sup[3]]+[list_sup[4]]]
        i+=1
    print('--------------------')
    # return 0
    j = 0
    for supp_num,list_date in dict_sup_pro.items():
        supp_num_convert = convert_to_input_sql(num_list=[supp_num])
        # To classify Check_column
        for date in list_date:
            df_splited = df_raw[(df_raw['VENDOR_NUMBER'] == supp_num) & (df_raw['STARTDATE'] == date[0])& (df_raw['ENDDATE'] == date[1])& (df_raw['CLASSIFY_TYPE'] == date[2]) & (df_raw['CLASSIFY_CLAIM'] == date[3])]   
            print(df_splited)
            if str(supp_num_convert) in list_ven_stop_trading:
                df_splited['CHECK_COLUMN'] = 'VENDOR STOP TRADING'
            elif df_splited['CLASSIFY_TYPE'].unique()[0] == 'CLAIM':
                df_splited['CHECK_COLUMN'] = 'CLAIM'
            elif df_splited['SUM_AMT_GAP'].unique()[0] < 100 :
                df_splited['CHECK_COLUMN'] = 'ELI_EXCLUDE < 100'
            else:
                if df_splited[['ITEMID','REBATE_ENTITLEMENT_NUM']].drop_duplicates().count()[0] == df_splited['ITEMID'].drop_duplicates().count():
                    startgap = df_splited['STARTDATE'].unique()[0]
                    endgap = df_splited['ENDDATE'].unique()[0]
                    classify_claim = df_splited['CLASSIFY_CLAIM'].unique()[0]
                    check_brandid  = df_splited[['BRANDID']].drop_duplicates().values.tolist()
                    check_brandid_final = []
                    for brandid in check_brandid:
                        check_brandid_final.append(brandid[0])
                    list_brandid = ",".join(check_brandid_final)
                    list_brandid = str(list_brandid)
                    df_splited['CHECK_BRANDID'] =  list_brandid
                    df_splited['COUNT_BRANDID']  = len(check_brandid_final)  
                    itemlist = df_splited['ITEMID'].drop_duplicates().values.tolist()
                    itemlist_convert = convert_to_input_function(itemlist)
                    uom_filter = df_splited['MULTIPLIER_NUM'].drop_duplicates().values.tolist()
                    uom_convert = convert_to_input_function(uom_filter)
                    brand_convert = convert_to_input_function(check_brandid_final)
                    check_cd = connect_sql(cursor = cursor,file_sql = file_sql_cd_check_again,item_code=itemlist_convert, var_1 =startgap,var_2 =endgap, var_3 = brand_convert,var_4 = uom_convert, var_5 = classify_claim)
                    if check_cd['CHECK_GAP'].unique()[0] == None:
                        df_splited['CHECK_COLUMN'] = 'TO QA'
                        df_splited['FINALSTARTGAP'] = df_splited['STARTDATE'] 
                        df_splited['FINALENDGAP'] = df_splited['ENDDATE'] 
                    else :
                        if check_cd['CHECK_GAP'].unique()[0] == 'NOGAP':
                            # df_splited['CHECK_COLUMN'] = 'PAID IN CD, CHECK CD AGAIN'
                            item_list =  df_splited[['ITEMID','BRANDID','MULTIPLIER_NUM','CLASSIFY_TYPE','CLASSIFY_CLAIM']].drop_duplicates().values.tolist()
                            i = 0
                            for item in item_list:
                                df_splited_2 = df_splited[(df_splited['ITEMID'] == item[0]) & (df_splited['BRANDID'] == item[1]) & (df_splited['MULTIPLIER_NUM'] == item[2])]
                                check_cd_2 = connect_sql(cursor = cursor,file_sql = file_sql_cd_check_again,item_code=item[0], var_1 =startgap,var_2 =endgap, var_3 = item[1],var_4 = str(item[2]), var_5 = classify_claim)
                                if check_cd_2['CHECK_GAP'].unique()[0] == None:
                                    df_splited_2['CHECK_COLUMN'] = 'TO QA'
                                    df_splited_2['FINALSTARTGAP'] = df_splited['STARTDATE'] 
                                    df_splited_2['FINALENDGAP'] = df_splited['ENDDATE']
                                    # df_splited_2
                                else:
                                    if check_cd_2['CHECK_GAP'].unique()[0] == 'NOGAP':
                                        df_splited_2['CHECK_COLUMN'] = 'PAID IN CD'
                                    else :
                                        df_splited_2['FINALSTARTGAP'] = check_cd_2['NEWSTARTGAP'].unique()[0]
                                        df_splited_2['FINALENDGAP'] = check_cd_2['NEWENDGAP'].unique()[0]
                                        df_splited_2['CHECK_COLUMN'] = 'PAID PARTITIALY IN CD, CHECK NEWGAP WITH NEWSTARTDATE AND NEWENDDATE'
                                if i  == 0:
                                    df_splited_part = df_splited_2
                                else:
                                    df_splited_part = pd.concat([df_splited_part, df_splited_2], ignore_index=True)
                                i+= 1
                            df_splited = df_splited_part
                        else:
                            df_splited['FINALSTARTGAP'] = check_cd['NEWSTARTGAP'].unique()[0]
                            df_splited['FINALENDGAP'] = check_cd['NEWENDGAP'].unique()[0]
                            df_splited['CHECK_COLUMN'] = 'PAID PARTITIALY IN CD, CHECK NEWGAP WITH NEWSTARTDATE AND NEWENDDATE'
                else:
                    df_splited['CHECK_COLUMN'] = 'CHECK AGAIN'
            if j  == 0:
                    df_raw_check = df_splited
            else:
                    df_raw_check = pd.concat([df_raw_check, df_splited], ignore_index=True)
            j+=1
    #Export checklist
    df_raw_check.to_excel(path_check_list,index=False)
    return 0
    ###############################################################
    error_list =[]
    time_start = datetime.datetime.now()
    # Filter df_splited with condition , keep TO QA and PRGX
    df_raw_filter = df_raw_check[(df_raw_check['CHECK_COLUMN'] == 'TO QA') | (df_raw_check['CHECK_COLUMN'] == 'PAID PARTITIALY IN CD, CHECK NEWGAP WITH NEWSTARTDATE AND NEWENDDATE')] 
    df_unique_supp_filter = df_raw_filter[['VENDOR_NUMBER','FINALSTARTGAP','FINALENDGAP','CLASSIFY_AMOUNT','CHECK_BRANDID','COUNT_BRANDID']].drop_duplicates().values.tolist()
    # Create dictionary with supp_num key and list of promo_ids after filter conditions, keep check again and to QA
    dict_sup_pro_filter = {}
    j=0

    #Get vendor_number + value(lessthan20k) to key of dictionary
    for list_sup in df_unique_supp_filter:
        if j == 0:
            dict_sup_pro_filter[list_sup[0]+list_sup[3]] = [list_sup]
        else:
            if list_sup[0]+list_sup[3] in dict_sup_pro_filter.keys():
                dict_sup_pro_filter[list_sup[0]+list_sup[3]].append(list_sup)
            else:
                dict_sup_pro_filter[list_sup[0]+list_sup[3]]  = [list_sup]
        j+=1

    #Split 2 brandid to each brandid/sheet, keep 3brandid -> LL
    for key,value_list in dict_sup_pro_filter.items():
        list_split = []
        for value in value_list:
            if value[5] == 2:
                split_all = value[4].split(',')
                for split in split_all:
                    list_split.append([value[0]]+[value[1]]+[value[2]]+[value[3]]+[split]+[value[5]])
            else:
                list_split.append(value)
        dict_sup_pro_filter[key] = list_split
    print(dict_sup_pro_filter)

    check_list_promo_index = 1
    for supp_classify_amt,date_list in dict_sup_pro_filter.items():
        index_promo=1
        # summary_index = 1
        summary_index_list = []
        for date in date_list:
            print('-------------------------------------------------------------------------------------------------------------------------------------')
            print(date)
            check_list_promo = []
            if date[5] != 3:
                df_splited_filter = df_raw_filter[(df_raw_filter['VENDOR_NUMBER'] == date[0]) & (df_raw_filter['FINALSTARTGAP'] == date[1])& (df_raw_filter['FINALENDGAP'] == date[2])& (df_raw_filter['CLASSIFY_AMOUNT'] == date[3])& (df_raw_filter['BRANDID'] == date[4]) & (df_raw_filter['COUNT_BRANDID'] == date[5])]
            else:
                df_splited_filter = df_raw_filter[(df_raw_filter['VENDOR_NUMBER'] == date[0]) & (df_raw_filter['FINALSTARTGAP'] == date[1])& (df_raw_filter['FINALENDGAP'] == date[2])& (df_raw_filter['CLASSIFY_AMOUNT'] == date[3])& (df_raw_filter['CHECK_BRANDID'] == date[4]) & (df_raw_filter['COUNT_BRANDID'] == date[5])]
            # get some important variable
            if df_splited_filter.empty :
                continue
            else:
                pass
            supp_num_list_final,item_list_final,item_input_function = get_info(df_splited = df_splited_filter)
            supp_desc = df_splited_filter['VENDOR_NAME'].unique()[0].replace("/","")
            print('supp_desc',supp_desc)
            try:
                paf_loc = df_splited_filter['PAF_LINK_SUGGEST'].unique()[0]
            except Exception :
                paf_loc = '0'
            try:
                email_loc = df_splited_filter['EMAIL_SUGGEST'].unique()[0]
            except Exception :
                email_loc = '0'
            vendor_num = df_splited_filter['VENDOR_NUMBER'].unique()[0] 
            clm_start = df_splited_filter['FINALSTARTGAP'].unique()[0]
            clm_end = df_splited_filter['FINALENDGAP'].unique()[0]
            classify_claim = df_splited_filter['CLASSIFY_CLAIM'].unique()[0]
            #lessthan20k
            amount = date[3]
            # brandid 
            if date[5] == 3 :
                brandid_merged = date[5]
            else:
                brandid_merged = df_splited_filter['BRANDID'].unique()[0]
            # uom
            check_uom = df_splited[['MULTIPLIER_NUM']].drop_duplicates().values.tolist()
            check_uom_final = []
            for uom in check_uom:
                check_uom_final.append(str(int(uom[0])))
            uom = ",".join(check_uom_final)
            #CATEGORY
            if date[5] == 3 :
                category_list = df_splited_filter[['CATEGORY_ID']].drop_duplicates().values.tolist()
                for category_element in category_list:
                    if 'LL' in category_element[0]:
                        category = category_element[0]
                        break
            else:
                category = df_splited_filter['CATEGORY_ID'].unique()[0]
            #create path for excel and path_xlsb for excel
            path_export_final = path_export+'CL_SCAN_'+supp_desc+'_Analyst_date'+amount+'.xlsx'
            path_export_final_xlsb = path_export+'CL_SCAN_'+supp_desc+'_Analyst_date_'+str(vendor_num)+amount+'.xlsb'
            create_worksheet(index_promo=str(index_promo),path_export_final=path_export_final)
            df_splited_filter_2 = df_splited_filter[['ITEMID','BRANDID','MULTIPLIER_NUM','PROMO_PRICE_SUGGEST','SCAN_RATE_SUGGEST']]
            df_splited_filter_2= df_splited_filter_2.drop_duplicates()
            df_splited_filter_2 = df_splited_filter_2.groupby(['ITEMID','BRANDID','MULTIPLIER_NUM']).agg({'PROMO_PRICE_SUGGEST':'max','SCAN_RATE_SUGGEST':'max'}).reset_index()
            print(df_splited_filter_2)

            # print(df_splited_filter_2)
            item_list_dict = df_splited_filter_2.set_index(['ITEMID','BRANDID','MULTIPLIER_NUM'])[['PROMO_PRICE_SUGGEST','SCAN_RATE_SUGGEST']].to_dict('index')
            for key,value in item_list_dict.items():
                item_list_dict[key] = [item_list_dict[key]['PROMO_PRICE_SUGGEST']] + [item_list_dict[key]['SCAN_RATE_SUGGEST']] 
            # To create excel file
            summary_index_list.append(index_promo)
            # df_sales = df_sales_data(cursor = cursor,item_list_dict_gsted = item_list_dict ,start_date = clm_start,end_date = clm_end)
            df_sales,list_data_sales,list_remove_sales = df_sales_data( cursor = cursor, file_sql= file_sql_summ , item_list_dict_gsted =item_list_dict,start_date = clm_start,end_date =clm_end)
            list_data_state,list_remove_state =  product_state_summary(df_sales = df_sales)
            list_data_product ,list_remove_product = product_summary(df_sales = df_sales)
            dict_data_dept = {'df':category,'cell_export':'C4'}
            dict_data_paf_loc = {'df':paf_loc,'cell_export':'R7'}
            dict_data_email_loc = {'df':email_loc,'cell_export':'S7'}
            dict_data_claim_number = {'df':str(index_promo),'cell_export':'B4'}
            # notes = 'Item "&TEXTJOIN(", ",1,UNIQUE(TOCOL(TotalSum[Product Number]&" - "&TotalSum[Product Name],1,1)))&"was on a promotion during this time. As per previous promotion, at the same RRP, the vendor would support a scan rate as below. "&"However according to our records, no funding has been charged. Please see sales data and email evidence for more information."'
            if classify_claim == 'EVD':
                dict_data_notes = {'df': 'EDV' ,'cell_export':'Q7'}
            else:
                dict_data_notes = {'df': 'PROMO' ,'cell_export':'Q7'}
            list_data = list_data_sales + list_data_state + list_data_product + [dict_data_dept] + [dict_data_paf_loc] + [dict_data_email_loc] + [dict_data_claim_number] + [dict_data_notes]
            list_remove = list_remove_sales + list_remove_state + list_remove_product
            #  Fill sheet Complete Daily Sales Data
            writer_excel(data = list_data, remove = list_remove,number_sheet= str(index_promo),path_export_final=path_export_final)  
            index_promo+=1
            check_list_promo_index += 1 
            with xw.App(visible=False) as app:
                print('check_list_promo_index')
                if check_list_promo_index == 2:
                    wb = app.books.open('check_list_promo.xlsx')
                else:
                    wb = app.books.open(path_check_list_promo)
                wb_sheet = wb.sheets['Sheet1']
                check_list_promo = [vendor_num] + [supp_desc]  + [clm_start] + [clm_end]+ [brandid_merged] + [uom] +['Done'] 
                print(check_list_promo)
                wb_sheet.range(f'A{check_list_promo_index}').value =  check_list_promo
                wb.save(path_check_list_promo)
        # Fill sheet Vendor Summary
        fill_summary_sheet(supp_desc = supp_desc,summary_index_list= summary_index_list,path_export_final=path_export_final,vendor_num = vendor_num)         
        remove_sheet_change_xlsb(sheet_name = 'template',path_export_final=path_export_final ,path_export_final_xlsb = path_export_final_xlsb)  
        print('-------------------------------------------------------------------------------------------------------------------------------------')
    print(datetime.datetime.now() - time_start)


if __name__ == '__main__':
    main()
    