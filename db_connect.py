import pandas as pd
import pyodbc


con_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=\\EP-FS-001\FileSharing\FINANCE\Accounting workings\EP Accounting Invoices V1.0.accdb;'
conn = pyodbc.connect(con_string)
cursor = conn.cursor()
cursor.execute('SELECT * FROM counterparties')


rows = []
for row in cursor.fetchall():
    rows.append([row[0], row[3]])

df_access = pd.DataFrame(rows, columns=['tax_code','mapping'])