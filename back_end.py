import pandas as pd
import pyodbc
import xml.etree.ElementTree as et


def backend(xlsx_path, xml_path):
    global xml_error, xlsx_error
    try:
        with open('db_path.txt', 'r') as f:
            dbpath = f.read()
        con_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=[' + dbpath + '];'
        conn = pyodbc.connect(con_string)
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM counterparties')


        rows = []
        for row in cursor.fetchall():
            rows.append([row[0], row[3]])

        df_access = pd.DataFrame(rows, columns=['tax_code','mapping'])
    except Exception as db_error:
        print(db_error)

    try:
        tree = et.parse(xml_path)

        root = tree.getroot()

        invoice_serial = []
        invoice_number = []
        status = []
        recorded_id = []
        issue_date = []
        delivery_date = []
        payment_date = []
        tax_code = []
        invoice_amount = []
        invoice_description = []
        standard_cost_id = []
        users = []
        comment = []

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}Series"):
            invoice_serial.append(elm.text)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}Number"):
            invoice_number.append(elm.text)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}Series"):
            status.append('Դուրս գրված')
            recorded_id.append('Ոչ')
            users.append('Անուն Ազգանուն')
            invoice_description.append('Գրեք')
            payment_date.append(0)
            standard_cost_id.append(0)
            comment.append(0)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}SubmissionDate"):
            issue_date.append(elm.text)

        if invoice_serial[0] == 'Բ':
            for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}DeliveryDate"):
                delivery_date.append(elm.text)
        else:
            for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}SupplyDate"):
                delivery_date.append(elm.text)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}SupplierInfo"):
            elm = elm.find(".//{http://www.taxservice.am/tp3/invoice/definitions}TIN")
            tax_code.append(elm.text)

        for elm in root.findall(".//{http://www.taxservice.am/tp3/invoice/definitions}Total"):
            elm = elm.find(".//{http://www.taxservice.am/tp3/invoice/definitions}TotalPrice")
            invoice_amount.append(elm.text)

        output = {
            'invoice_serial': invoice_serial,
            'invoice_number': invoice_number,
            'status': status,
            'recorded_id': recorded_id,
            'issue_date': issue_date,
            'delivery_date': delivery_date,
            'payment_date': payment_date,
            'tax_code': tax_code,
            'invoice_description': invoice_description,
            'standard_cost_id': standard_cost_id,
            'users': users,
            'comment': comment
        }

        pd.set_option('display.max_columns', None)
        df = pd.DataFrame(output)

        df['Invoice ID'] = df['invoice_serial'] + df['invoice_number'].astype(str)
        df['issue_date'] = df['issue_date'].str[:10]
        df['issue_date'] = pd.to_datetime(df['issue_date']).dt.strftime('%d/%m/%Y')
        df['delivery_date'] = df['delivery_date'].str[:10]
        df['delivery_date'] = pd.to_datetime(df['delivery_date']).dt.strftime('%d/%m/%Y')
        df['payment_date'].replace([0], '', inplace=True)
        df['standard_cost_id'].replace([0], '', inplace=True)
        df['comment'].replace([0], '', inplace=True)

        df = df[[
            'Invoice ID',
            'status',
            'recorded_id',
            'issue_date',
            'delivery_date',
            'payment_date',
            'tax_code',
            'invoice_description',
            'standard_cost_id',
            'users',
            'comment'
        ]]

        df_join = pd.merge(df, df_access, on='tax_code', how='left')

        writer = pd.ExcelWriter(xlsx_path)
        df_join.to_excel(writer)
        writer.save()

    except Exception as xml_error:
        print(xml_error)