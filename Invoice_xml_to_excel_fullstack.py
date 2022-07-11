import PySimpleGUI as sg
import xml.etree.ElementTree as et
import pandas as pd
import sys
from db_connect import df_access

sg.theme("green")
sg.set_options(font="Arial", element_text_color="white", text_color="white")

mainlayout = [
    [sg.Text("Խնդրում եմ ներմուծել eInvoice-ից արտահանված xml ֆայլը")],
    [sg.InputText(key="-XML_PATH-", disabled=True), sg.FileBrowse(file_types=[('XML Files', '*.xml')])],
    [sg.Text("Խնդրում եմ նշել թե որտեղ արտահանել excel ֆայլը")],
    [sg.InputText(key="-XLSX_PATH-", disabled=True), sg.FolderBrowse()],
    [sg.Button("Արտահանել Excel-ը", key="-RUN-", size=(53, 1))]
]

MainWindow = sg.Window("XML Invoice to Excel", mainlayout)

while True:
    event, values = MainWindow.read()
    if event in (sg.WIN_CLOSED, 'Quit'):
        MainWindow.close()
        sys.exit()

    elif event == '-RUN-' and values['-XML_PATH-'] == "" or values['-XLSX_PATH-'] == "":
        sg.popup("Զգուշացում", "Խնդրում եմ լրացրեք բոլոր դաշտերը")
        continue

    elif event == '-RUN-':
        xml_path = values['-XML_PATH-']
        xlsx_path = values['-XLSX_PATH-']
        xlsx_path = xlsx_path + '/invoices.xlsx'
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

            sg.popup("Կատարված է", "Invoice-ները բարեհաջող արտահանվել են")
            continue
        except Exception:
            sg.popup("Անհաջող փորձ", "Invoice-ը չստացվեց արտահանել։ Ստուգեք արդյոք օգտագործում եք ճիշտ xml ֆայլ։"
                                     "Խնդրի պատճառը չհայտնաբերելու պարագայում դիմեք Հայկին:")
            continue