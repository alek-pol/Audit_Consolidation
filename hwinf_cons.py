import os
import glob
import lxml
import openpyxl
from bs4 import BeautifulSoup
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment

src_path = 'c:\Development\Audit\src\*'
dst_path = 'c:\Development\Audit\\'
dst_file = 'Test1.xlsx'

min_win10_build = 16299  # 1709
sncp = {
    'Raw': {'SheetName': 'Raw', 'CurRow': 1, 'CurCol': 0},
    'List': {'SheetName': 'List', 'CurRow': 0, 'CurCol': 0},
    'Consolidation': {'SheetName': 'Consolidation', 'CurRow': 0, 'CurCol': 0}
}
sncp_lock = list(sncp.keys())
list_col = []


def gen_list_col():
    l_col = []
    for f_ch in ['', 'A', 'B']:
        for n_ch in range(26):
            l_col.append((f_ch + chr(65 + n_ch)))
    return l_col


def gen_cr(list_n):
    cur_cr = list_col[sncp[list_n]['CurCol']] + str(sncp[list_n]['CurRow'])
    sncp[list_n]['CurCol'] += 1
    return cur_cr

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
        sncp['List']['CurRow'] += 1
        if key not in sncp_lock:
            make_tab_body(new_wb['List'], 'A%s' % sncp['List']['CurRow'], sncp[key]['SheetName'], 'Main_center')
            hyperlink ='=HYPERLINK("#%s!A1", "%s")' % (sncp[key]['SheetName'], key)
            make_tab_body(new_wb['List'], 'B%s' % sncp['List']['CurRow'], hyperlink, 'Main_hyperlink')


def make_raw_topic():


    make_main_topic(new_wb['Raw'], 'A1', 'Свод данных', 'Topic_main')


    make_tab_topic(new_wb['Raw'], 'A3', 'Наименование организации', 'Topic_tab', 50)
    make_tab_topic(new_wb['Raw'], 'B3', 'Рабочее место', 'Topic_tab', 20)
    make_tab_topic(new_wb['Raw'], 'C3', 'Название АРМ', 'Topic_tab', 10)
    make_tab_topic(new_wb['Raw'], 'D3', 'Операционная система', 'Topic_tab', 30)
    make_tab_topic(new_wb['Raw'], 'E3', 'Кол-во ядер', 'Topic_tab', 8)
    make_tab_topic(new_wb['Raw'], 'F3', 'Кол-во логич-х проц-в', 'Topic_tab', 9)
    make_tab_topic(new_wb['Raw'], 'G3', 'Процессор', 'Topic_tab', 20)
    make_tab_topic(new_wb['Raw'], 'H3', 'Сокет', 'Topic_tab', 10)
    make_tab_topic(new_wb['Raw'], 'I3', 'L3 - Кэш', 'Topic_tab', 10)
    make_tab_topic(new_wb['Raw'], 'J3', 'Модель мат.платы', 'Topic_tab', 15)
    make_tab_topic(new_wb['Raw'], 'K3', 'Год произв-ва', 'Topic_tab', 10)





    make_tab_topic(new_wb['Raw'], 'BX3', 'Пользователь', 'Topic_tab', 15)

    sncp['Raw']['CurRow'] = 4


def make_org_topic(list_n, org_name):
    make_main_topic(new_wb[list_n], 'A1', org_name, 'Topic_main')
    make_tab_topic(new_wb[list_n], 'A3', '№', 'Topic_tab', 5)
    make_tab_topic(new_wb[list_n], 'B3', 'Рабочее место', 'Topic_tab', 18)
    make_tab_topic(new_wb[list_n], 'C3', 'Операционная система', 'Topic_tab', 20)
    make_tab_topic(new_wb[list_n], 'D3', 'Необходимо обновление ОС', 'Topic_tab', 13)
    make_tab_topic(new_wb[list_n], 'E3', 'Процессор', 'Topic_tab', 15)
    make_tab_topic(new_wb[list_n], 'F3', 'Кол-во ядер', 'Topic_tab', 8)
    make_tab_topic(new_wb[list_n], 'G3', 'Сокет', 'Topic_tab', 8)


    sncp[org_name]['CurRow'] += 1







def scan_hwi_htm(src_hwi_path):
    list_path = src_hwi_path.split('\\')
    llp = len(list_path)
    hwi_htm = open(src_hwi_path, 'r')
    soup = BeautifulSoup(hwi_htm, 'lxml')  # Parse the HTML as a string
    tables = soup.find_all('table')
    len_tables = len(tables)

    org = list_path[llp - 3]
    workplace = list_path[llp - 2]

    info_base = {
        "Computer Name:": "",
        "Operating System:": "",
        "Current User Name:": ""
    }
    info_cpu_1 = {
        "Number Of Processor Cores:": "",
        "Number Of Logical Processors:": ""
    }
    info_cpu_2 = {
        "CPU Brand Name:": "",
        "CPU Code Name:": "",
        "CPU Technology:": "",
        "CPU Platform:": "",
        "L3 Cache:": ""
    }

    info_mb = {
        "Motherboard Model:": "",
        "BIOS Date:": ""
    }

    def make_headers_list(tables_n):
        headers = []
        for t in range(len(tables_n)):
            tmp = None
            for col in tables_n[t].find_all("td", class_="dt"):
                tmp = col.get_text()
            if tmp:
                headers.append(tmp)
            else:
                headers.append('')
        return headers

    def scan_value(table_n, f_param):
        flag_p = False
        current_p = ''
        for row in table_n.find_all('tr'):
            for column in row.find_all('td'):
                tmp_s = column.get_text()
                if tmp_s in f_param and flag_p == False:
                    current_p = tmp_s
                    flag_p = True
                elif flag_p:
                    flag_p = False
                    f_param[current_p] = tmp_s.rstrip().lstrip()

    def check_os(os_n):
        if os:
            if os_n[2] not in ['10']:
                return 'Да'
            else:
                if 'Home' in os_n:
                    return 'Да'
                elif int((os_n[int(os_n.index('Build')) + 1]).split('.')[0]) < min_win10_build:
                    return 'Да'
                else:
                    return 'Нет'

    def check_date(date_m):
        date_m = int(date_m.split('/')[2])
        if int(date_m) < 100:
            date_m += 2000
        return date_m








    # ====================
    header_list = make_headers_list(tables)
    org_sn = sncp[org]['SheetName']
    raw_cr = sncp['Raw']['CurRow']
    org_cr = sncp[org]['CurRow']
    # ------------------------------------------------------------------------------------
    scan_value(tables[3], info_base)

    index_cpu = header_list.index('Central Processor(s)')
    scan_value(tables[index_cpu + 1], info_cpu_1)
    scan_value(tables[index_cpu + 3], info_cpu_2)

    # Motherboard info
    index_mb = header_list.index('Motherboard')
    scan_value(tables[index_mb + 1], info_mb)







    # entry Raw list
    make_tab_body(new_wb['Raw'], 'A%s' % raw_cr, org, 'Main_left')
    make_tab_body(new_wb['Raw'], 'B%s' % raw_cr, workplace, 'Main_left')
    make_tab_body(new_wb['Raw'], 'C%s' % raw_cr, info_base["Computer Name:"], 'Main_left')
    make_tab_body(new_wb['Raw'], 'D%s' % raw_cr, info_base["Operating System:"], 'Main_left')
    make_tab_body(new_wb['Raw'], 'E%s' % raw_cr, info_cpu_1["Number Of Processor Cores:"], 'Main_center')
    make_tab_body(new_wb['Raw'], 'F%s' % raw_cr, info_cpu_1["Number Of Logical Processors:"], 'Main_center')
    make_tab_body(new_wb['Raw'], 'G%s' % raw_cr, info_cpu_2["CPU Brand Name:"], 'Main_center')
    make_tab_body(new_wb['Raw'], 'H%s' % raw_cr, info_cpu_2["CPU Platform:"], 'Main_center')
    make_tab_body(new_wb['Raw'], 'I%s' % raw_cr, info_cpu_2["L3 Cache:"], 'Main_center')
    make_tab_body(new_wb['Raw'], 'J%s' % raw_cr, info_mb["Motherboard Model:"], 'Main_left')
    make_tab_body(new_wb['Raw'], 'K%s' % raw_cr, check_date(info_mb["BIOS Date:"]), 'Main_center')


    # entry Org List
    make_tab_body(new_wb[org_sn], 'A%s' % org_cr, org_cr - 3, 'Main_left')
    make_tab_body(new_wb[org_sn], 'B%s' % org_cr, workplace, 'Main_left')
    make_tab_body(new_wb[org_sn], 'C%s' % org_cr, info_base["Operating System:"], 'Main_left')
    make_tab_body(new_wb[org_sn], 'D%s' % org_cr, check_os(info_base["Operating System:"].split(' ')), 'Main_center')
    make_tab_body(new_wb[org_sn], 'E%s' % org_cr, info_cpu_2["CPU Brand Name:"], 'Main_left')
    make_tab_body(new_wb[org_sn], 'F%s' % org_cr, info_cpu_1["Number Of Processor Cores:"], 'Main_left')
    make_tab_body(new_wb[org_sn], 'G%s' % org_cr, info_cpu_2["L3 Cache:"], 'Main_left')



    # -----------------------------------------------
    # entry CPU info















    # entry misk info

    make_tab_body(new_wb['Raw'], 'BX%s' % sncp['Raw']['CurRow'], info_base["Current User Name:"], 'Main_left')





    sncp['Raw']['CurRow'] += 1
    sncp[org]['CurRow'] += 1


# Рейтинг тех. оснащенности
#
#
#
#
#
# ===================================================
if __name__ == '__main__':


    # Create new excel file
    new_wb = openpyxl.Workbook()
    init_style()
    list_col = gen_list_col()





    org_path = glob.glob(src_path)
    for i in range(len(org_path)):
        org_name = org_path[i].split('\\')[-1]
        if org_name not in sncp:
            sncp[org_name] = {'SheetName': str(i + 1), 'CurRow': 1}

    for key in reversed(list(sncp.keys())):
        new_wb.create_sheet(title=sncp[key]['SheetName'], index=0)
        if key not in sncp_lock:
            make_org_topic(sncp[key]['SheetName'], key)
            sncp[key]['CurRow'] += 2

    make_list()  # Список листов с гиперссылкой
    make_raw_topic()

    for i in range(len(org_path)):
        print('----------------------------------------------------------------------------')
        print(org_path[i])
        workplaces = glob.glob(org_path[i]+'\*')
        hwi_htm_files = glob.glob(org_path[i]+'\*\*.htm*')
        print(len(workplaces), len(hwi_htm_files))
        org_name = org_path[i].split('\\')[-1]
        j = 0
        for src_tmp_path in hwi_htm_files:
            #print(src_tmp_path)

            scan_hwi_htm(src_tmp_path)
            j += 1
            if j > 3:
                break
        new_wb.save(dst_path + dst_file)
        if i > 6:
            break


















# ========================================

    new_wb.save(dst_path + dst_file)
    print('Save file')



