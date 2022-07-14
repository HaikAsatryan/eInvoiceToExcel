con_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=[\\EP-FS-001\FileSharing\FINANCE\Accounting workings\EP Accounting Invoices V1.0.accdb];'

with open('db_path.txt', 'r') as f:
    dbpath = f.read()

con_string2 = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=[' + dbpath + '];'

print(con_string==con_string2)