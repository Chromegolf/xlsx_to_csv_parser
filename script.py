import pandas as pd
import openpyxl

wb = openpyxl.load_workbook('export.xlsx')
ws = wb.active
max_col = ws.max_column
max_row = ws.max_row


def concat_precondition(row_value):
    values = []
    del values[:]
    for row in ws.iter_cols(min_col=2, max_col=2, min_row=2, max_row=row_value):
        for cell in row:
            if cell.value is None:
                pass
            else:
                values.append(str(cell.value))

    ws[f'B2'].value = ';'.join(values)
    for row in range(2, row_value):
        ws[f'B{row+1}'].value = None
    wb.save('mode.xlsx')


def concat_step(row_value):
    values = []
    del values[:]
    for row in ws.iter_cols(min_col=3, max_col=3, min_row=2, max_row=row_value):
        for cell in row:
            if cell.value is None:
                pass
            else:
                values.append(str(cell.value))

    ws[f'C2'].value = ';'.join(values)
    for row in range(2, row_value):
        ws[f'C{row+1}'].value = None
    wb.save('mode.xlsx')


def export_to_csv():
    df = pd.read_excel(r'mode.xlsx')
    df.to_csv('export.csv', index=None, header=True)
    df = pd.DataFrame(pd.read_csv("export.csv"))


if __name__ == '__main__':
    concat_precondition(max_row)
    concat_step(max_row)
    export_to_csv()