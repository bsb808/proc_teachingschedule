import openpyxl 
planner='MAE_Planner_AY24.xlsx'
wb = openpyxl.load_workbook(planner)
ws = wb.active

rows = []

for row in ws.iter_rows(
        min_row = 10, max_row=11, min_col=1, max_col=ws.max_column,
        values_only=True):
    rows.append(row)

for row in rows:
    for cell in row:
        if not (cell is None):
            n = cell.find('[')
            m = cell.find(']')
            if ((n > 0) and (m > 0)):
                print(cell[n+1:m])
                                     
    
