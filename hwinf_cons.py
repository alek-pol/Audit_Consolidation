import os
import glob
import openpyxl
from bs4 import BeautifulSoup
from collections import OrderedDict
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment

src_path = 'c:\Development\Audit\src\*'

dst_path = 'c:\Development\Audit\\'
dst_file = 'Test1.xlsx'

sncp = {
    'Raw': {'SheetName': 'RAW', 'CurrentRow': 1},
    'List': {'SheetName': 'List', 'CurrentRow': 0},
    'Consolidation': {'SheetName': 'Consolidation', 'CurrentRow': 0}
}
sncp_lock = list(sncp.keys())

def add_style(workbook, new_style, **kwargs):
    if new_style not in workbook.named_styles:
        n_style = NamedStyle(name=new_style)
        for key, value in kwargs.items():
            setattr(n_style, key, value)
        workbook.add_named_style(n_style)


def init_style():
    b_side = Side(style='thin', color='000000')

    add_style(new_wb, 'Topic_main',
              font=Font(name='Calibri', size=12, bold=True),
              alignment=Alignment(vertical='center'),
              )

    add_style(new_wb, 'Topic_tab',
              font=Font(name='Calibri', size=9, bold=True),
              alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side)
              )

    add_style(new_wb, 'Main_center',
              font=Font(name='Calibri', size=9),
              alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side)
              )

    add_style(new_wb, 'Main_left',
              font=Font(name='Calibri', size=9),
              alignment=Alignment(vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side)
              )

    add_style(new_wb, 'Main_hyperlink',
              font=Font(name='Calibri', size=9, color='00008B'),
              alignment=Alignment(vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side)
              )


def make_main_topic(sheet_n, col_row, value_n, style_n):
    sheet_n[col_row] = value_n
    sheet_n[col_row].style = style_n
    sheet_n.row_dimensions[1].height = 25


def make_tab_topic(sheet_n, col_row, value_n, style_n, dimension_n):
    sheet_n[col_row] = value_n
    sheet_n[col_row].style = style_n
    sheet_n.column_dimensions[col_row[0]].width = dimension_n


def make_tab_body(sheet_n, col_row, value_n, style_n):
    sheet_n[col_row] = value_n
    sheet_n[col_row].style = style_n


def make_list():
    make_main_topic(new_wb['List'], 'A1', 'Список листов', 'Topic_main')
    make_tab_topic(new_wb['List'], 'A3', '№ Листа', 'Topic_tab', 6)
    make_tab_topic(new_wb['List'], 'B3', 'Наименование организации', 'Topic_tab', 80)
    for key in sncp.keys():
        sncp['List']['CurrentRow'] += 1
        if key not in sncp_lock:
            print('A%s' % sncp['List']['CurrentRow'], sncp[key]['SheetName'], key)
            make_tab_body(new_wb['List'], 'A%s' % sncp['List']['CurrentRow'], sncp[key]['SheetName'], 'Main_center')
            hyperlink ='=HYPERLINK("#%s!A1", "%s")' % (sncp[key]['SheetName'], key)
            make_tab_body(new_wb['List'], 'B%s' % sncp['List']['CurrentRow'], hyperlink, 'Main_hyperlink')






# ===================================================
if __name__ == '__main__':


    # Create new excel file
    new_wb = openpyxl.Workbook()
    init_style()




    org_path = glob.glob(src_path)
    for i in range(len(org_path)):
        org_name = org_path[i].split('\\')[-1]
        if org_name not in sncp:
            sncp[org_name] = {'SheetName': str(i + 1), 'CurrentRow': 1}

    for key in reversed(list(sncp.keys())):
        new_wb.create_sheet(title=sncp[key]['SheetName'], index=0)
        if key not in sncp_lock:
            make_main_topic(new_wb[sncp[key]['SheetName']], 'A1', key, 'Topic_main')
            sncp[key]['CurrentRow'] += 1

    make_list()






















# ========================================

    new_wb.save(dst_path + dst_file)
    print('Save file')



