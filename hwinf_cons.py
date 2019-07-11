import openpyxl
from bs4 import BeautifulSoup
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment


def add_style(workbook, new_style, **kwargs):
    if new_style not in workbook.named_styles:
        n_style = NamedStyle(name=new_style)
        for key, value in kwargs.items():
            setattr(n_style, key, value)
        workbook.add_named_style(n_style)


# ===================================================
if __name__ == '__main__':
    print('--')


# ---------------- Test ------------------------
    # Create new excel file
    new_wb = openpyxl.Workbook()
    new_wb.create_sheet(title='RAW', index=0)
    sheet = new_wb['RAW']



    b_side = Side(style='thin', color='000000')
    add_style(new_wb, 'Header',
              font=Font(name='Calibri', size=9, bold=True),
              alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side)
              )

    sheet['A2'].style = 'Header'
    sheet['A2'] = 'Test'

    new_wb.save('c:\Development\Audit\Test1.xlsx')
    print('[Ok]')
