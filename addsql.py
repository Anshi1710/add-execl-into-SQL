from sqlalchemy import create_engine # pip install SQLAlchemy
from sqlalchemy.engine import URL
import pypyodbc # pip install pypyodbc
import pandas as pd # pip install pandas

SERVER_NAME = 'IDEAPAD\SQLEXPRESS (SQL Server 16.0.1000 - Ideapad\anshi)'
DATABASE_NAME = 'anshidb'
TABLE_NAME = 'demo_table'

excel_file = r'D:\internship\5_plasticCardPrint.csv'

connection_string = f"""
    DRIVER={{SQL Server}};
    SERVER={SERVER_NAME};
    DATABASE={DATABASE_NAME};
    Trusted_Connection=yes;
"""
connection_url = URL.create('mssql+pyodbc', query={'odbc_connect': connection_string})
enigne = create_engine(connection_url, module=pypyodbc)

excel_file = pd.read_excel(excel_file, sheet_name=None)
for sheet_name, df_data in excel_file.items():
    print(f'Loading worksheet {sheet_name}...')
    # {'fail', 'replace', 'append'}
    df_data.to_sql(TABLE_NAME, enigne, if_exists='append', index=False)