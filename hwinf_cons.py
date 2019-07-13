import os
import glob
import lxml
import openpyxl
from bs4 import BeautifulSoup
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill

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

out_data_raw = {
    'Organithation': {
        'name': 'Наименование организации', 'b_style': 'Main_left', 'width': 30, 'value': ''},
    'Workplace': {
        'name': 'Рабочее место', 'b_style': 'Main_left', 'width': 16, 'value': ''},
    'Computer Name:': {
        'name': 'Название АРМ', 'b_style': 'Main_left', 'width': 10, 'value': ''},
    'Status': {
        'name': 'Статус', 'b_style': 'Main_left', 'width': 10, 'value': ''},
    'Operating System:': {
        'name': 'Операционная система', 'b_style': 'Main_center', 'width': 20, 'value': ''},
    'OS update': {
        'name': 'Необходимость обновить ОС', 'b_style': 'Main_center', 'width': 15, 'value': ''},
    'Number Of Processor Cores:': {
        'name': 'Кол-во ядер', 'b_style': 'Main_center', 'width': 8, 'value': ''},
    'Number Of Logical Processors:': {
        'name': 'Кол-во логич-х проц-в', 'b_style': 'Main_center', 'width': 9, 'value': ''},
    'CPU Brand Name:': {
        'name': 'Процессор', 'b_style': 'Main_center', 'width': 20, 'value': ''},
    'CPU Platform:': {
        'name': 'Сокет', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'L3 Cache:': {
        'name': 'L3 - Кэш', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Motherboard Model:': {
        'name': 'Модель мат.платы', 'b_style': 'Main_center', 'width': 15, 'value': ''},
    'Data': {
        'name': 'Год произв-ва', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Total Memory Size:': {
        'name': 'Опре. памяти всего', 'b_style': 'Main_center', 'width': 10, 'value': ''}

    # '': {'name': '', 'b_style': '', 'width': '', 'value': []},
}


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

    add_style(new_wb, 'Main_center_red',
              font=Font(name='Calibri', size=9),
              alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side),
              fill=PatternFill(fgColor="FFA07A", patternType='solid')
              )

    add_style(new_wb, 'Main_center_yellow',
              font=Font(name='Calibri', size=9),
              alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side),
              fill=PatternFill(fgColor="FFE4B5", patternType='solid')
              )

    add_style(new_wb, 'Main_center_green',
              font=Font(name='Calibri', size=9),
              alignment=Alignment(horizontal='center', vertical='center', wrap_text=True),
              border=Border(left=b_side, top=b_side, right=b_side, bottom=b_side),
              fill=PatternFill(fgColor="98FB98", patternType='solid')
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
    make_main_topic(new_wb['Raw'], 'A1', 'Общий свод данных', 'Topic_main')
    sncp['Raw']['CurRow'] = 3
    sncp['Raw']['CurCol'] = 0
    for k, v in out_data_raw.items():
        make_tab_topic(new_wb['Raw'], gen_cr('Raw'), v['name'], 'Topic_tab', v['width'])
    sncp['Raw']['CurRow'] += 1


def out_raw_body():
    sncp['Raw']['CurCol'] = 0
    if out_data_raw['Organithation']['value'] != '':
        for k, v in out_data_raw.items():
            make_tab_body(new_wb['Raw'], gen_cr('Raw'), v['value'], v['b_style'])
            v['value'] = ''

        sncp['Raw']['CurRow'] += 1


def make_org_topic(list_n, org_name):
    make_main_topic(new_wb[list_n], 'A1', org_name, 'Topic_main')
    sncp[org_name]['CurRow'] = 3
    sncp[org_name]['CurCol'] = 0
    make_tab_topic(new_wb[list_n], gen_cr(org_name), '№', 'Topic_tab', 5)
    make_tab_topic(new_wb[list_n], gen_cr(org_name), 'Рабочее место', 'Topic_tab', 16)
    make_tab_topic(new_wb[list_n], gen_cr(org_name), 'Операционная система', 'Topic_tab', 20)
    make_tab_topic(new_wb[list_n], gen_cr(org_name), 'Необходимо обновление ОС', 'Topic_tab', 13)
    make_tab_topic(new_wb[list_n], gen_cr(org_name), 'Процессор', 'Topic_tab', 15)
    make_tab_topic(new_wb[list_n], gen_cr(org_name), 'Кол-во ядер', 'Topic_tab', 8)
    make_tab_topic(new_wb[list_n], gen_cr(org_name), 'Сокет', 'Topic_tab', 8)


    sncp[org_name]['CurRow'] = 4







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
        "Current User Name:": "",
        'Manufacturer:': "",
        'Case Type:': ""


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


    socket_stat = {
        '0': {}


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

    def check_date(date_m):
        date_m = int(date_m.split('/')[2])
        if int(date_m) < 100:
            date_m += 2000
        if date_m <= 2010:
            return {'value': date_m, 'age': 'old', 'color': 'Main_center_red'}
        elif date_m >= 2011 and date_m <= 2013:
            return {'value': date_m, 'age': 'middle', 'color': 'Main_center_yellow'}
        else:
            return {'value': date_m, 'age': 'new', 'color': 'Main_center_green'}

    def check_os(os_n):
        if os_n:
            if os_n[2] not in ['10']:
                return 'Да', 'Main_center_red'
            else:
                if 'Home' in os_n:
                    return 'Да', 'Main_center_red'
                elif int((os_n[int(os_n.index('Build')) + 1]).split('.')[0]) < min_win10_build:
                    return 'Да', 'Main_center_red'
                else:
                    return 'Нет', 'Main_center_green'
        return 'проверка', 'Main_center_yellow'



    # ====================
    header_list = make_headers_list(tables)

    # ------------------------------------------------------------------------------------
    # Base info
    scan_value(tables[3], info_base)
    oper_sys = (check_os(info_base["Operating System:"].split(' ')))

    index_base = header_list.index('System Enclosure')
    scan_value(tables[index_base + 1], info_base)
    print(info_base['Case Type:'])




    # CPU info
    index_cpu = header_list.index('Central Processor(s)')
    scan_value(tables[index_cpu + 1], info_cpu_1)
    scan_value(tables[index_cpu + 3], info_cpu_2)

    # Motherboard info
    index_mb = header_list.index('Motherboard')
    scan_value(tables[index_mb + 1], info_mb)
    info_date = check_date(info_mb["BIOS Date:"])




    #if date_status == 1:
     #   status = ['на списание', '']
    #else:
     #   status = ['', '']



    # entry Raw list

    out_data_raw['Organithation']['value'] = org
    out_data_raw['Workplace']['value'] = workplace
    out_data_raw['Computer Name:']['value'] = info_base["Computer Name:"]
    #out_data_raw['Status']['value'] = status[0]

    out_data_raw['Operating System:']['value'] = info_base["Operating System:"]
    out_data_raw['OS update']['value'] = oper_sys[0]
    out_data_raw['OS update']['b_style'] = oper_sys[1]

    out_data_raw['Number Of Processor Cores:']['value'] = info_cpu_1["Number Of Processor Cores:"]
    out_data_raw['Number Of Logical Processors:']['value'] = info_cpu_1["Number Of Logical Processors:"]
    out_data_raw['CPU Brand Name:']['value'] = info_cpu_2["CPU Brand Name:"]
    out_data_raw['CPU Platform:']['value'] = info_cpu_2["CPU Platform:"]
    out_data_raw['L3 Cache:']['value'] = info_cpu_2["L3 Cache:"]

    out_data_raw['Motherboard Model:']['value'] = info_mb["Motherboard Model:"]
    out_data_raw['Data']['value'] = info_date['value']
    out_data_raw['Data']['b_style'] = info_date['color']
    #out_data_raw['']['value'] = ''




    #make_tab_body(new_wb['Raw'], gen_cr('Raw'), , 'Main_left')
    #make_tab_body(new_wb['Raw'], gen_cr('Raw'), , )
    #make_tab_body(new_wb['Raw'], gen_cr('Raw'), , 'Main_left')



    # entry Org List
    org_sn = sncp[org]['SheetName']
    sncp[org]['CurCol'] = 0
    make_tab_body(new_wb[org_sn], gen_cr(org), 1, 'Main_left')
    make_tab_body(new_wb[org_sn], gen_cr(org), workplace, 'Main_left')
    make_tab_body(new_wb[org_sn], gen_cr(org), info_base["Operating System:"], 'Main_left')
    make_tab_body(new_wb[org_sn], gen_cr(org), oper_sys[0], 'Main_center')
    make_tab_body(new_wb[org_sn], gen_cr(org), info_cpu_2["CPU Brand Name:"], 'Main_left')
    make_tab_body(new_wb[org_sn], gen_cr(org), info_cpu_1["Number Of Processor Cores:"], 'Main_center')
    make_tab_body(new_wb[org_sn], gen_cr(org), info_cpu_2["CPU Platform:"], 'Main_center')

    #make_tab_body(new_wb[org_sn], gen_cr(org), info_cpu_2["L3 Cache:"], 'Main_left')








    # entry misk info
    #make_tab_body(new_wb['Raw'], 'BX%s' % sncp['Raw']['CurRow'], info_base["Current User Name:"], 'Main_left')

    #sncp['Raw']['CurRow'] += 1
    sncp[org]['CurRow'] += 1






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

            out_raw_body()

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



