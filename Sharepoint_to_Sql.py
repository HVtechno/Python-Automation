import sharepy
import pandas as pd
import pyodbc
from fast_to_sql import fast_to_sql as fts
import os

#GET ALL ENV VARIABLES
site_url = os.getenv('sharepy_site_url')
username = os.getenv('username')
password = os.getenv('password')
file_path = os.getenv('sharepy_file_path')
DB_DRIVER =  os.getenv('DB_DRIVER')
DB_SERVER = os.getenv('DB_SERVER')
DB_DATABASE = os.getenv('DB_DATABASE')
DB_USER = os.getenv('DB_USER')
DB_PASSWORD = os.getenv('DB_PASSWORD')

#CONNECTION TO SQL DB(QA)
Hconn = "Driver=" + "{" + DB_DRIVER + "};" + "Server=" + DB_SERVER + ";" + "Database=" + DB_DATABASE + ";" + "uid=" + DB_USERNAME + ";" + "pwd=" + DB_PASSWORD + ";" + "Trusted_Connection = yes;"
cnxn = pyodbc.connect(Hconn)

try:
    #AUTHENTICATE & CREATE SHAREPY OBJECT
    s = sharepy.connect(site_url, username, password)
    if not hasattr(s,'cookie'):
        print("Authentication Success!!")
        #DOWNLOAD FILE & SAVE IT TO MEMORY
        response = s.getfile(file_path)
        print("Response received & file downloaded to memory!!")
        #READ EXCEL & CONVERT TO DATAFRAME(SHEETS:'Documents')
        Sharepointdata = pd.read_excel(open('./DOCbycustomer.xlsx','rb'),sheet_name='Documents')
        HV_sharedata = pd.DataFrame(Sharepointdata)
        #CONDITIONAL FILTERING (HARD CODED IN SHIPTO FOR TO EXCLUDE '29744324')
        if (HV_sharedata['Shipto'].isin([29744324]).any()):
            HV_sharedata = HV_sharedata.loc[HV_sharedata['Shipto'] != 29744324]
        if (HV_sharedata['Export Zone'].str.contains(str("Not found")).any()):
            HV_sharedata = HV_sharedata.loc[HV_sharedata['Export Zone'].str.contains(str("Not found"))==False]
        if (HV_sharedata['Status'].str.contains(str("Blocked")).any()):
            HV_sharedata = HV_sharedata.loc[HV_sharedata['Status'].str.contains(str("Blocked"))==False]
        #RENAME COLUMNS AS PER DATABASE COLUMN AND READY TO UPLOAD
        HV_sharedata.columns = ['Export Zone','Country','Country Code','Sales Org','Soldto','Sold to name','Shipto','Ship to name','Payer','Payer name','Status','Sh.Cond code','Shipping','Prepayment','KPI',
                                'Customer type','Sending System','DOCS','BL','INVOICE','PL','COO','COA/CAT','COH','COI','Other documents','Special Requirements','Annual docs','Hard copy docs','Documentation Contact',
                                'DHL address','Tips & tricks','Documentation owner','Back up 1','Back up 2','note']
        mapping = {HV_sharedata.columns[0] : 'export_zone',HV_sharedata.columns[1] : 'country',HV_sharedata.columns[2]:'country_code',HV_sharedata.columns[3]:'sales_organization',HV_sharedata.columns[4]:'sold_to_number',HV_sharedata.columns[5]:'sold_to_name',HV_sharedata.columns[6]:'ship_to_number',HV_sharedata.columns[7]:'ship_to_name',HV_sharedata.columns[8]:'payer_number',HV_sharedata.columns[9]:'payer_name',HV_sharedata.columns[10]:'status',
                   HV_sharedata.columns[11]:'shipping_conditions',HV_sharedata.columns[12]:'shipping',HV_sharedata.columns[13]:'prepayment',HV_sharedata.columns[14]:'kpi',HV_sharedata.columns[15]:'customer_type',HV_sharedata.columns[16]:'sending_system',HV_sharedata.columns[17]:'docs',HV_sharedata.columns[18]:'bill_of_lading',HV_sharedata.columns[19]:'invoice',HV_sharedata.columns[20]:'packing_list',HV_sharedata.columns[21]:'coo',
                   HV_sharedata.columns[22]:'coa_cat',HV_sharedata.columns[23]:'coh',HV_sharedata.columns[24]:'coi',HV_sharedata.columns[25]:'other_documents',HV_sharedata.columns[26]:'special_requirements',HV_sharedata.columns[27]:'annual_docs',HV_sharedata.columns[28]:'hard_copy_docs',HV_sharedata.columns[29]:'documentation_contact',HV_sharedata.columns[30]:'dhl_address',HV_sharedata.columns[31]:'tips_tricks',HV_sharedata.columns[32]:'documentation_owner',
                   HV_sharedata.columns[33]:'back_up1',HV_sharedata.columns[34]:'back_up2',HV_sharedata.columns[35]:'note'}
        Final_HV_sharedata = HV_sharedata.rename(columns=mapping)
        #UPLOAD DATAFRAME TO DATABASE
        with cnxn.cursor() as cursor:
            try:
                print("Database Connected successfully!!!")
                DEL_share = '''Delete FROM [dbo].[Customer_Table]'''
                cursor.execute(DEL_share)
                print('Records deleted successfully!!') 
                create_statement = fts.fast_to_sql(Final_HV_sharedata, "Customer_Table", cnxn, if_exists="append", temp=False)
                print('Sharepoint data uploaded to SQL successfully!!') 
                cnxn.commit()
                cnxn.close()
            except (pyodbc.Error, pyodbc.Warning) as e:
                print(e)
                print("Sorry, there is an issue with the processing of data or Database Connection not established!!!")
        #UNLINK SP DOWNLOADED FILE FROM MEMORY
        HV_SP_data = './DOCbycustomer.xlsx'
        if os.path.isfile(HV_SP_data):
            os.unlink(HV_SP_data)
        else:
            print("Error: file (DOCbycustomer.xlsx) not found")
    #ABORT THE PROCESS IF AUTHENTICATION FAILED
    else:
        print("Authentication failed!!")
except Exception as error:
    print(error)