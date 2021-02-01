from report_readers import customs_report
from invoice import Invoice, InvoiceWriter

records = customs_report('input_files/customs_report_sample.xls')

invoices = {}
for record in records:
    try:
        invoices[record[3]].add_record(record)
    except KeyError:
        invoices.update({record[3]:Invoice(record, '210131')})

writer = InvoiceWriter('input_files/driver_list.xlsx', 'input_files/invoice_template.xlsx')

for invoice in invoices.values():
    writer.write_invoice(invoice, 'input_files')