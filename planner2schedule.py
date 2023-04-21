import openpyxl 
planner='MAE_Planner_AY24.xlsx'
wb = openpyxl.load_workbook(planner)
ws = wb.active

vals = []
for val in ws.iter_rows(
        min_row = 10, max_row=10, min_col=1, max_col=ws.max_column,
        values_only=True):
    vals.append(val)


