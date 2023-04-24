from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import pandas as pd
from io import BytesIO
import pyodbc
from fast_to_sql import fast_to_sql as fts
import time
import os

#GET ALL .ENV Variables
site_url = os.getenv('site_url')
username = os.getenv('username')
password = os.getenv('password')
folder_url = os.getenv('folder_url')
DB_DRIVER =  os.getenv('DB_DRIVER')
DB_SERVER = os.getenv('DB_SERVER')
DB_DATABASE = os.getenv('DB_DATABASE')
DB_USER = os.getenv('DB_USER')
DB_PASSWORD = os.getenv('DB_PASSWORD')

# AUTHENTICATE SHAREPOINT USING CREDENTIALS
ctx_auth = AuthenticationContext(site_url)
ctx_auth.acquire_token_for_user(username, password)
ctx = ClientContext(site_url, ctx_auth)

#READ EXCEL FILE
df_excel_loop = pd.read_excel('Excel_Loop.xlsx')

# GET ALL FILES FROM SHAREPOINT FOLDER
folder = ctx.web.get_folder_by_server_relative_url(folder_url)
files = folder.files
ctx.load(files)
ctx.execute_query()

# LOOP THROUGH EACH EXCEL
for file in files:
    file_name = file.properties["Name"]
    if file_name.endswith(".xlsx"):
        # OPEN THE EXCEL FILES AS BINARY
        response = File.open_binary(ctx, file.serverRelativeUrl)
        #SEND THE RESPONSE AND GET THE FILES
        if response.status_code == 200 and response.headers.get("Content-Type") == "application/octet-stream":
            file_content = response.content
            # PROCESS THE FILE CONTENT USING OPENPYXL or PANDAS AS BELOW
            excel_data = pd.read_excel(BytesIO(file_content))
            #CONNECTION TO SQL DB
            Hconn = "Driver=" + "{" + DB_DRIVER + "};" + "Server=" + DB_SERVER + ";" + "Database=" + DB_DATABASE + ";" + "uid=" + DB_USERNAME + ";" + "pwd=" + DB_PASSWORD + ";" + "Trusted_Connection = yes;"
            cnxn = pyodbc.connect(Hconn)
            #CHECK IF CONNECTION EXIST
            try:
                cnxn.execute("SELECT 1")
                print(f"Database is connected and {file_name} ready to upload")
                create_statement = fts.fast_to_sql(excel_data, "SP_Loop_Data", cnxn, if_exists="append", temp=False)
                cnxn.commit()
                new_row = {'File_Name': file_name,'Status': 'Data Loaded to DB Successfully'}
                df_excel_loop = df_excel_loop.append(new_row, ignore_index=True)
            #CATCH ERROR IF CONNECTION NOT, EXIST
            except pyodbc.Error:
                print(f"Database is not connected and {file_name} cannot be loaded")
                new_row = {'File_Name': file_name,'Status': 'Data not Loaded to DB due to connection issues.'}
                df_excel_loop = df_excel_loop.append(new_row, ignore_index=True)
            finally:
                #FINALLY SAVE THE FILE
                df_excel_loop.to_excel('Excel_Loop.xlsx', index=False)
                cnxn.close()
                time.sleep(5)
        #CATCH THE ERROR IF FILE NOT RECEIVED FROM THE RESPONSE API
        else:
            print(f"Error downloading file {file_name}")
            print(f"Status code: {response.status_code}")
            print(f"Content type: {response.headers.get('Content-Type')}")