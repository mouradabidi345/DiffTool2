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
# engine = sqlalchemy.create_engine("mssql+pyodbc://" + "mourada:" + "J+!_b7jHm`+(5" +'"'+"!s" +"@InfoTrax_Prod") #working
engine = sqlalchemy.create_engine("mssql+pyodbc://" + "aseauser:" + "@S34Pr0d"+"@ASEA_PROD") #working
# engine = sqlalchemy.create_engine("mssql+pyodbc://aseauser:@S34Pr0d@ASEA_PROD")
# engine = sqlalchemy.create_engine("mssql+pyodbc://aseareadonly:C2UCjlqte9mV6}z@InfoTrax_Prod")
# df = pd.read_sql_query('SELECT * FROM InfoTrax_Prod.dbo.[Tbl_Orders_Header]', engine) # working
# df1 = pd.read_sql_query('SELECT * FROM InfoTrax_Prod.dbo.[TBL_ORDERS_PRODUCTS]', engine) # working

# df = pd.read_sql_query("""SELECT * FROM InfoTrax_Prod.dbo.[TBL_ORDERS_PRODUCTS] p inner join InfoTrax_Prod.dbo.[Tbl_Orders_Header] oh on p.ORDER_NUMBER = oh.LegacyNumber WHERE OH.ORDERDATE >= '2021/01/01'""", engine)
df = pd.read_sql_query("""SELECT distinct p.*
FROM ASEA_PROD.dbo.TBL_ORDERS_PRODUCTS p
inner join ASEA_PROD.dbo.Tbl_Orders_Header oh                                                                                          
on p.ORDER_NUMBER = oh.LegacyNumber                                       
WHERE OH.ORDERDATE >= '2021/01/01'""", engine)



#df = pd.read_sql_query('SELECT * FROM ASEA_PROD.dbo.[TBL_DISTRIBUTOR]', engine) #working
df.columns = map(str.upper, df.columns)
# df1.columns = map(str.upper, df1.columns)
#new
# dfmerged = df1.merge(df[df.ORDERDATE >= '2018-01-01'], left_on= 'ORDER_NUMBER', right_on= 'LEGACYNUMBER')

# df1 = df[(df.CREATEDDATE > '2018-01-01' ) & (df.CREATEDDATE < start_date) & (df.UPDATEDDATE <= start_date)] 

# df1 = df.query('CREATEDDATE > "2015-01-10" and CREATEDDATE < start_date and UPDATEDDATE <= start_date', inplace = True)

ctx = snowflake.connector.connect(
          user='MOURADABIDI',
          password='Ma@07842032',
          account='ba62849.east-us-2.azure',
          warehouse= 'COMPUTE_MACHINE',
          database='DB_ASEA_REPORTS',
          schema='PUBLIC')    



# cur = ctx.cursor()
# Table =  'Tbl_Orders_Header'   
# # # Execute a statement that will generate a result set.
# warehouse= 'COMPUTE_MACHINE'
# database='DB_ASEA_REPORTS'
# schema='PUBLIC'
# if warehouse:
#     cur.execute(f'use warehouse {warehouse};')
#     cur.execute(f'select * from {database}.{schema}.{Table};')
# # Fetch the result set from the cursor and deliver it as the Pandas DataFrame.
# snowflakedf = cur.fetch_pandas_all()


# Table2 =  'TBL_ORDERS_PRODUCTS'   
# # # Execute a statement that will generate a result set.
# warehouse= 'COMPUTE_MACHINE'
# database='DB_ASEA_REPORTS'
# schema='PUBLIC'
# if warehouse:
#     cur.execute(f'use warehouse {warehouse};')
#     cur.execute(f'select * from {database}.{schema}.{Table2};')
# # Fetch the result set from the cursor and deliver it as the Pandas DataFrame.
# snowflakedf1 = cur.fetch_pandas_all()

# snowflakedfmerged = snowflakedf1.merge(snowflakedf[snowflakedf.ORDERDATE >= '2018-01-01'], left_on= 'ORDER_NUMBER', right_on= 'LEGACYNUMBER')



# compare = datacompy.Compare(
# df1,
# snowflakedf1,
# join_columns= 'LEGACYNUMBER')
# compare.matches(ignore_extra_columns=False) 
# print(compare.report())
# sqldatabase = 'InfoTrax_Prod'
#     #sqldatabase = 'ASEA_PROD'
# Today = datetime.datetime.today()
# outlook = client.Dispatch('Outlook.Application')
# message = outlook.Createitem(0)
# message.Display()
# message.To = 'mabidi@aseaglobal.com'
# message.Subject = 'SQL COMPARE With Window Time '  + sqldatabase +'.'+ Table + ' ' + ' as of ' + ' ' + str(Today)
# message.Body = compare.report()
# message.Save()
# message.Send()

# cur.close()




def  SnowflakeQA(Table, Primary_key):
    cur = ctx.cursor()
    
# # Execute a statement that will generate a result set.
    warehouse= 'COMPUTE_MACHINE'
    database='DB_ASEA_REPORTS'
    schema='PUBLIC'
    if warehouse:
        cur.execute(f'use warehouse {warehouse};')
    # cur.execute("""SELECT * FROM DB_ASEA_REPORTS.PUBLIC.TBL_ORDERS_PRODUCTS p inner join DB_ASEA_REPORTS.PUBLIC.Tbl_Orders_Header oh on p.ORDER_NUMBER = oh.LegacyNumber WHERE OH.ORDERDATE >= '01/01/2021';""")
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
