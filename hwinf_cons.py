import os
import glob
import openpyxl
from bs4 import BeautifulSoup
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment

sncp = {} # sheet name and current positon

def add_style(workbook, new_style, **kwargs):
    if new_style not in workbook.named_styles:
        n_style = NamedStyle(name=new_style)
        for key, value in kwargs.items():
            setattr(n_style, key, value)
        workbook.add_named_style(n_style)


def init_style():
    b_side = Side(style='thin', color='000000')
    add_style(new_wb, 'Header_0',
              font=Font(name='Calibri', size=11, bold=True),
              alignment=Alignment(vertical='center'),
              )

    add_style(new_wb, 'Header',
              font=Font(name='Calibri', size=9, bold=True),
              alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side)
              )

    add_style(new_wb, 'Main_a_center',
              font=Font(name='Calibri', size=9),
              alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side)
              )

    add_style(new_wb, 'Main_a_left',
              font=Font(name='Calibri', size=9),
              alignment=Alignment(vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side)
              )














# ===================================================
if __name__ == '__main__':

    src_path = 'c:\Development\Audit\src\*'

    # Create new excel file
    new_wb = openpyxl.Workbook()
    init_style()

    # Create new Sheet - Raw
    sncp = {'RAW': {'SheetName': 'RAW', 'CurrentPosition': '1'}}
    if new_wb.sheetnames:
        new_wb[new_wb.sheetnames[0]].title = 'Raw'
        print()

    new_wb.create_sheet(title='List', index=0)

    org_path = glob.glob(src_path)
    for i in range(len(org_path)):
        org_name = org_path[i].split('\\')[-1]
        if org_name not in sncp:
            sncp[org_name] = {'SheetName': str(i + 1), 'CurrentPosition': '2'}

            new_wb.create_sheet(title=sncp[org_name]['SheetName'], index=0)
            sheet = new_wb[str(i+1)]
            sheet['A1'] = org_name
            sheet['A1'].style = 'Header_0'

            # print(sncp[org_name]['SheetName'], org_name)












# ========================================
    new_wb.save('c:\Development\Audit\Test1.xlsx')
    print('Save file')



