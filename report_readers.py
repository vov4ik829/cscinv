from workbook import CellOutOfRange, load_workbook as load_ro_workbook

class ReportTypeError(Exception):
    pass


def customs_report(filename:str):
    '''Reads customs report

    Args:
        filename (str): name of the xls or xlsx file

    Returns:
    list: list of form: [[container_id, type, full/empty, car_id],]
    '''
    wb = load_ro_workbook(filename)
    ws = wb.worksheets[0]
    if [
        ws['A1'].value.strip(), ws['C1'].value.strip(), 
        ws['D1'].value.strip(), ws['E1'].value.strip()
        ]!=[
            'კონტეინერის ნომერი','გემის ვიზიტის №', 
            'დატვირთული/ცარიელი', 'მანქანის ნომერი №'
            ]:
        raise ReportTypeError
    row = 2
    while ws['{}{}'.format('A', row)].value is not None:

        cont_id = ws['{}{}'.format('A',row)].value
        cont_type = ws['{}{}'.format('B',row)].value
        conv = {'იმპორტი':'full', 'ექსპორტი':'full', 'ცარიელი':'empty'}
        cargo = conv[ws['{}{}'.format('D', row)].value]
        car_id = ws['{}{}'.format('E', row)].value.split('/')[0]

        yield [cont_id, cont_type, cargo, car_id]
        row+=1
        try: # Checking if end reached 
            ws['A{}'.format(row)].value = ws['A{}'.format(row)].value
        except CellOutOfRange:
            break
    wb.close()

