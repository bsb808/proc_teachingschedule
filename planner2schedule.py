import openpyxl 

class Course:
    def __init__(self,number, title):
        self.number = number
        self.title = title
        self.currics = []

class Section:
    def __init__(self, qtr, number, title, is_dl):
        self.number = number
        self.title = title
        self.qtr = qtr
        self.is_dl = is_dl
        self.currics = []

def is_in_sections(sections, s):
    '''  Is the sectin in a list of sections? '''
    for sec in sections:
        if ((sec.number == s.number) and (sec.qtr == s.qtr)):
            return True
    return False

def add_section(sections, qtr, number, title, is_dl, curric):                  
        # Create a section - counts on a sections list at is modified in place
        s = Section(qtr, number, title, is_dl)
        if not is_in_sections(sections, s):
            sections.append(s)
            
class Instructor:
    def __init__(self, name):
        self.name = name
        self.fa_re = []
        self.fa_dl = []
        self.wi_re = []
        self.wi_dl = []
        self.sp_re = []
        self.sp_dl = []
        self.su_re = []
        self.su_dl = []
        self.sections = []

    def add_section(self, qtr, number, title, is_dl, curric):                  
        # Create a section
        s = Section(qtr, number, title, is_dl)
        if not is_in_sections(self.sections, s):
            self.sections.append(s)
                      
                      
def is_in_instructors(instructors, i):
    '''  Is the sectin in a list of sections? '''
    for instructor in instructors:
        if (instructor.name == i.name):
            return True
    return False

        
        
planner='MAE_Planner_AY24.xlsx'
wb = openpyxl.load_workbook(planner)
ws = wb.active

rows = []

# Row with DL designation
dl = 5
for row in ws.iter_rows(
        min_row = dl, max_row = dl,
        min_col=1, max_col=ws.max_column,
        values_only=True):
    dl_row = row

# Merged dl_row
mdl_row = []
for c, col_index in zip(dl_row,range(ws.max_column)):
    val = c
    if c is None:
        for crange in ws.merged_cells:
            clo,rlo,chi,rhi = crange.bounds
            top_value = ws.cell(rlo,clo).value
            if (rlo<=dl and dl<=rhi
                and clo<=col_index and col_index<=chi):
                val = top_value
                break
        # Replace None with blank string
        val = ""
    mdl_row.append(val)
    
# Specify the 4 rows for the AY we are interested in.
fall_row = 9
summer_row = 12

for row in ws.iter_rows(
        min_row = fall_row, max_row = summer_row,
        min_col=1, max_col=ws.max_column,
        values_only=True):
    rows.append(row)

# Keys are instructor names, values are a list of Sections
instructors = dict()

qtrs = ['fall', 'winter', 'spring', 'summer']

for row, qtr in zip(rows, qtrs):
    for cell, dl_cell in zip(row, mdl_row):
        if not (cell is None):
            n = cell.find('[')
            m = cell.find(']')
            if ((n > 0) and (m > 0)):                    
                iname = cell[n+1:m]
                part = cell[:n-1]
                ss = part.splitlines()
                if len(ss) < 1:
                    print("Error with cell: "+cell)
                    break
                cnum = ss[0]
                if len(ss) > 1:
                    ctitle = "".join(ss[1:])
                else:
                    ctitle = ""

                # DL or not?
                isdl = False
                if dl_cell.find('DL') > 0:
                    isdl = True
                instructors.setdefault(iname, [])
                add_section(instructors[iname], qtr, cnum, ctitle, isdl, [])


# Print
for k in instructors.keys():
    print(k, end=": ")
    sections = instructors[k]
    for qtr in qtrs:
        print(qtr, end=": ")
        for s in sections:
            if s.qtr == qtr:
                if s.is_dl:
                    print(s.number, end="(DL), ")
                else:
                    print(s.number, end=", ")                    

    print("")
    
                    
                                     
    
