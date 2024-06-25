import xlsxwriter
import json


my_format = []
with open("./cas-17022024.json", "r") as file:
    cas = json.loads(file.read())
    for folio in cas['data']['folios']:
        amc = folio['amc']
        folio_id = folio['folio']
        for scheme in folio['schemes'][0]['transactions']:
            amount=scheme['amount']
            date = scheme['date']
            nav = scheme['nav']
            type = scheme['type']
            units = scheme['units']
            my_format.append([amc, folio_id, amount, date, nav, type, units])






with xlsxwriter.Workbook('cas-17022024.xlsx') as workbook:
    worksheet = workbook.add_worksheet()
    worksheet.write_row(0, 0, ["AMC", "Folio ID", "Amount", "Date", "NAV", "Type", "Units"])
    row=1 # 0 is header row
    col=0
    for index, entry in enumerate(my_format):
        worksheet.write_row(row + index, col, entry)


    
    