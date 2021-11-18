import pandas as pd
import xlrd
import sqlalchemy
import math
from sqlalchemy import create_engine
from pandas_profiling import ProfileReport
from snowflake.connector.pandas_tools import write_pandas
import snowflake.connector
import numpy as np
import win32com.client as client
import datetime
import datacompy

start_date = datetime.datetime.today() - datetime.timedelta(30)
start_date = start_date.strftime("%Y-%m-%d")
###########load the snowflake table
# engine = sqlalchemy.create_engine("mssql+pyodbc://" + "")
engine = sqlalchemy.create_engine("mssql+pyodbc://" + "") 


# df = pd.read_sql_query("""SELECT * FROM InfoTrax_Prod.dbo.[TBL_ORDERS_PRODUCTS] p inner join InfoTrax_Prod.dbo.[Tbl_Orders_Header] oh on p.ORDER_NUMBER = oh.LegacyNumber WHERE OH.ORDERDATE >= '2021/01/01'""", engine)
df = pd.read_sql_query("""SELECT distinct p.*
FROM ASEA_PROD.dbo.TBL_ORDERS_PRODUCTS p
inner join ASEA_PROD.dbo.Tbl_Orders_Header oh                                                                                          
on p.ORDER_NUMBER = oh.LegacyNumber                                       
WHERE OH.ORDERDATE >= '2021/01/01'""", engine)





ctx = snowflake.connector.connect(
          user='',
          password='',
          account='',
          warehouse= '',
          database='',
          schema='')    




def  SnowflakeQA(Table, Primary_key):
    cur = ctx.cursor()
    
# # Execute a statement that will generate a result set.
    warehouse= ''
    database=''
    schema=''
    if warehouse:
        cur.execute(f'use warehouse {warehouse};')
    
    cur.execute("""SELECT distinct p.*
    FROM DB_ASEA_REPORTS.PUBLIC.TBL_ORDERS_PRODUCTS p
    inner join DB_ASEA_REPORTS.PUBLIC.Tbl_Orders_Header oh                                                                                          
    on p.ORDER_NUMBER = oh.LegacyNumber                                       
    WHERE OH.ORDERDATE >= '01/01/2021';""")
# Fetch the result set from the cursor and deliver it as the Pandas DataFrame.
    snowflakedf = cur.fetch_pandas_all()
    


    compare = datacompy.Compare(
    df,
    snowflakedf,
    join_columns= Primary_key)
    compare.matches(ignore_extra_columns=False) 
    print(compare.report())
    # sqldatabase = 'InfoTrax_Prod'
    sqldatabase = 'ASEA_PROD'
    Today = datetime.datetime.today()
    outlook = client.Dispatch('Outlook.Application')
    message = outlook.Createitem(0)
    message.Display()
    message.To = 'mabidi@aseaglobal.com'
    message.Subject = 'SQL COMPARE With Window Time'  + sqldatabase +'.'+ Table + ' ' + ' as of ' + ' ' + str(Today)
    message.Body = compare.report()
    message.Save()
    message.Send()

    cur.close()

result = SnowflakeQA('TBL_ORDERS_PRODUCTS', 'DETAIL_ID')
