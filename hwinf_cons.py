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
    'Case Type:': {
        'name': 'Тип корпуса', 'b_style': 'Main_left', 'width': 10, 'value': ''},
    'Status': {
        'name': 'Статус', 'b_style': 'Main_left', 'width': 10, 'value': ''},
    'Operating System:': {
        'name': 'Операционная система', 'b_style': 'Main_left', 'width': 20, 'value': ''},
    'OS need update': {
        'name': 'Необходимо обновить ОС', 'b_style': 'Main_center', 'width': 11, 'value': ''},
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
        'name': 'Год произв-ва', 'b_style': 'Main_center', 'width': 7, 'value': ''},
    'Total Memory Size:': {
        'name': 'Опер. памяти всего', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Total Memory Size [MB]:': {
        'name': 'Опер. памяти всего (Mb)', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Maximum Supported Memory Clock:': {
        'name': 'Максимально поддерживаемая скороость', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Maximum Memory Size per Channel:': {
        'name': 'Максимум памяти на канал', 'b_style': 'Main_center', 'width': 10, 'value': ''},

    'Need upgrade memory': {
        'name': 'Необходимо увеличить память', 'b_style': 'Main_center', 'width': 11, 'value': ''},
    'Total Memory Count': {
        'name': 'Кол-во модулей', 'b_style': 'Main_center', 'width': 10, 'value': ''},


    'Module Size:0': {
        'name': 'Модуль 1 объем', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Memory Speed:0': {
        'name': 'Модуль 1 скорость', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Module Manufacturer:0': {
        'name': 'Модуль 1 производитель', 'b_style': 'Main_left', 'width': 13, 'value': ''},
    'Module Part Number:0': {
        'name': 'Модуль 1 PartNumber', 'b_style': 'Main_left', 'width': 10, 'value': ''},

    'Module Size:1': {
        'name': 'Модуль 2 объем', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Memory Speed:1': {
        'name': 'Модуль 2 скорость', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Module Manufacturer:1': {
        'name': 'Модуль 2 производитель', 'b_style': 'Main_left', 'width': 13, 'value': ''},
    'Module Part Number:1': {
        'name': 'Модуль 2 PartNumber', 'b_style': 'Main_left', 'width': 10, 'value': ''},

    'Module Size:2': {
        'name': 'Модуль 3 объем', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Memory Speed:2': {
        'name': 'Модуль 3 скорость', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Module Manufacturer:2': {
        'name': 'Модуль 3 производитель', 'b_style': 'Main_left', 'width': 13, 'value': ''},
    'Module Part Number:2': {
        'name': 'Модуль 3 PartNumber', 'b_style': 'Main_left', 'width': 10, 'value': ''},

    'Module Size:3': {
        'name': 'Модуль 4 объем', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Memory Speed:3': {
        'name': 'Модуль 4 скорость', 'b_style': 'Main_center', 'width': 10, 'value': ''},
    'Module Manufacturer:3': {
        'name': 'Модуль 4 производитель', 'b_style': 'Main_left', 'width': 13, 'value': ''},
    'Module Part Number:3': {
        'name': 'Модуль 4 PartNumber', 'b_style': 'Main_left', 'width': 10, 'value': ''}


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
        'Computer Name:': '',
        'Current User Name:': '',
        'Manufacturer:': '',
        'Case Type:': ''
    }
    info_os = {
        'Operating System:': '',
        'OS Version': '',
        'OS Category': '',
        'OS Build': 0,
        'OS Type': '',
        'OS need update': 'Да',
        'Update style': 'Main_center_red'
    }
    info_date = {
        'value': '',
        'age': '',
        'style': 'Main_center'
    }
    info_cpu = {
        'Number Of Processor Cores:': '',
        'Number Of Logical Processors:': '',
        'CPU Brand Name:': '',
        'CPU Code Name:': '',
        'CPU Technology:': '',
        'CPU Platform:': '',
        'L3 Cache:': '',
        'Maximum Memory Size per Channel:': ''
    }
    info_mb = {
        "Motherboard Model:": "",
        "BIOS Date:": ""
    }
    info_memory = {
        'Total Memory Size:': '',
        'Total Memory Size [MB]:': '',
        'Maximum Supported Memory Clock:': '',
        'Current Memory Clock:': '',
        'Current Timing (tCAS-tRCD-tRP-tRAS):': '',
        'Memory Runs At:': '',
        'Memory Type:': '',
        'Module Type:': '',
        'Need upgrade memory': '',
        'Upgrade style': 'Main_center',
        'Total Memory Count': 0
        }
    info_module = {
        'Module Density:': ['', '', '', ''],
        'Memory Speed:': ['', '', '', ''],
        'Module Manufacturer:': ['', '', '', ''],
        'Module Part Number:': ['', '', '', '']
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
        info_date['value'] = date_m
        if date_m <= 2010:
            info_date['age'] = 'old'
            info_date['style'] = 'Main_center_red'
        elif date_m >= 2011 and date_m <= 2013:
            info_date['age'] = 'middle'
            info_date['style'] = 'Main_center_yellow'
        else:
            info_date['age'] = 'new'
            info_date['style'] = 'Main_center_green'

    def check_os(os_n):
        if 'Microsoft' in os_n:
            build = os_n[os_n.index('Build') + 1].split('.')[0]
            if build:
                info_os['OS Build'] = int(build)
            if '(x64)' in os_n:
                info_os['OS Type'] = 'x64'
            if os_n[2] in ['10', '8.1', '8', '7', 'Vista', 'XP', '2000']:
                info_os['OS Version'] = os_n[2]
            elif os_n[2] == 'Server':
                info_os['OS Version'] = os_n[2] + ' ' + os_n[3]

            if info_os['OS Version'] == '10' and info_os['OS Build'] >= min_win10_build:
                info_os['OS need update'] = 'Нет'
                info_os['Update style'] = 'Main_center_green'

    def check_memory():
        info_memory['Total Memory Size [MB]:'] = int(info_memory['Total Memory Size [MB]:'])
        if info_memory['Total Memory Size [MB]:'] < 4096:
            info_memory['Need upgrade memory'] = 'Да'
            info_memory['Upgrade style'] = 'Main_center_red'
        else:
            info_memory['Need upgrade memory'] = 'Нет'

    def check_module():
        index_module = []
        index_memdev = []
        module_tmp = {
            'Module Density:': '',
            'Memory Type:': '',
            'Module Type:': '',
            'Memory Speed:': '',
            'Module Manufacturer:': '',
            'Module Part Number:': ''
            }
        memdev_tmp0 = {
            'Device Size:': '',  # Module Density:
            'Device Type:': '',  # Memory Type:
            'Device Form Factor:': '',   # Module Type:
            'Memory Speed:': '',  # Memory Speed:
            'Manufacturer:': '',  # Module Manufacturer:
            'Part Number:': ''  # Module Part Number:
        }

        def wr_type(im_name, im_value):
            if im_value != '' and info_memory[im_name] == '':
                info_memory[im_name] = im_value


        for t in range(header_list.index('Memory'), len_tables-50):
            if header_list[t].startswith('Row:'):
                index_module.append(t)

        if len(index_module) != 0:
            for im in range(len(index_module)):
                scan_value(tables[index_module[im] + 1], module_tmp)
                for mkey in info_module.keys():
                    info_module[mkey][im] = module_tmp[mkey]
                    module_tmp[mkey] = ''
                wr_type('Memory Type:', module_tmp['Memory Type:'])
                wr_type('Module Type:', module_tmp['Module Type:'])
            info_memory['Total Memory Count'] = len(index_module)
        else:
            print('--- ind 0 ------------------')
            s_start = header_list.index('Memory Devices')
            for t in range(s_start + 1, s_start + 15):
                if header_list[t] == 'Memory Device':
                    index_memdev.append(t)
            for im in range(len(index_memdev)):
                memdev_tmp = memdev_tmp0
                scan_value(tables[index_memdev[im] + 1], memdev_tmp)
                if memdev_tmp['Device Size:'] != '0 MBytes':
                    module_tmp['Module Density:'] = memdev_tmp['Device Size:']
                    module_tmp['Memory Speed:'] = memdev_tmp['Memory Speed:']
                    module_tmp['Module Manufacturer:'] = memdev_tmp['Manufacturer:']
                    module_tmp['Module Part Number:'] = memdev_tmp['Part Number:']
                    wr_type('Memory Type:', memdev_tmp['Device Type:'])
                    wr_type('Module Type:', memdev_tmp['Device Form Factor:'])
                    info_memory['Total Memory Count'] += 1






















    # ====================
    header_list = make_headers_list(tables)
    print(workplace)
    # ------------------------------------------------------------------------------------
    # Base info
    scan_value(tables[3], info_base)
    scan_value(tables[3], info_os)
    index_base = header_list.index('System Enclosure')
    scan_value(tables[index_base + 1], info_base)

    # Operating System info
    check_os(info_os["Operating System:"].split(' '))

    # CPU info
    index_cpu = header_list.index('Central Processor(s)')
    scan_value(tables[index_cpu + 1], info_cpu)
    scan_value(tables[index_cpu + 3], info_cpu)
    print('max ram', info_cpu['Maximum Memory Size per Channel:'])

    # Motherboard info
    index_mb = header_list.index('Motherboard')
    scan_value(tables[index_mb + 1], info_mb)
    check_date(info_mb["BIOS Date:"])

    # Scan Memory Device
    index_memory = header_list.index('Memory')
    scan_value(tables[index_memory + 1], info_memory)
    check_memory()
    #info_memory['Total Memory Count'] = header_list.count('Memory Device')
    check_module()



    #if date_status == 1:
     #   status = ['на списание', '']
    #else:
     #   status = ['', '']



    # entry Raw list
    out_data_raw['Organithation']['value'] = org
    out_data_raw['Workplace']['value'] = workplace
    out_data_raw['Computer Name:']['value'] = info_base['Computer Name:']
    out_data_raw['Case Type:']['value'] = info_base['Case Type:']
    #out_data_raw['Status']['value'] = status[0]

    out_data_raw['Operating System:']['value'] = info_os['Operating System:']
    out_data_raw['OS need update']['value'] = info_os['OS need update']
    out_data_raw['OS need update']['b_style'] = info_os['Update style']

    out_data_raw['Number Of Processor Cores:']['value'] = info_cpu['Number Of Processor Cores:']
    out_data_raw['Number Of Logical Processors:']['value'] = info_cpu['Number Of Logical Processors:']
    out_data_raw['CPU Brand Name:']['value'] = info_cpu['CPU Brand Name:']
    out_data_raw['CPU Platform:']['value'] = info_cpu['CPU Platform:']
    out_data_raw['L3 Cache:']['value'] = info_cpu['L3 Cache:']

    out_data_raw['Motherboard Model:']['value'] = info_mb['Motherboard Model:']
    out_data_raw['Data']['value'] = info_date['value']
    out_data_raw['Data']['b_style'] = info_date['style']

    out_data_raw['Total Memory Size:']['value'] = info_memory['Total Memory Size:']
    out_data_raw['Need upgrade memory']['value'] = info_memory['Need upgrade memory']
    out_data_raw['Need upgrade memory']['b_style'] = info_memory['Upgrade style']

    out_data_raw['Total Memory Size [MB]:']['value'] = info_memory['Total Memory Size [MB]:']
    out_data_raw['Maximum Supported Memory Clock:']['value'] = info_memory['Maximum Supported Memory Clock:']
    out_data_raw['Maximum Memory Size per Channel:']['value'] = info_cpu['Maximum Memory Size per Channel:']
    out_data_raw['Total Memory Count']['value'] = info_memory['Total Memory Count']




    #for i in range(info_memory['Total Memory Count']):
        #print('----', i, info_module['Module Size:'][i])
        #out_data_raw['Module Size:%i' % i]['value'] = info_module['Module Size:'][i]
        #out_data_raw['Memory Speed:%s' % i]['value'] = info_module['Memory Speed:'][i]
        #out_data_raw['Module Manufacturer:%s' % i]['value'] = info_module['Module Manufacturer:'][i]
        #out_data_raw['Module Part Number:%s' % i]['value'] = info_module['Module Part Number:'][i]






    #out_data_raw['']['value'] = ''
    #print(round(1.5))




    #make_tab_body(new_wb['Raw'], gen_cr('Raw'), , 'Main_left')
    #make_tab_body(new_wb['Raw'], gen_cr('Raw'), , )
    #make_tab_body(new_wb['Raw'], gen_cr('Raw'), , 'Main_left')



    # entry Org List
    #org_sn = sncp[org]['SheetName']
    #sncp[org]['CurCol'] = 0
    #make_tab_body(new_wb[org_sn], gen_cr(org), 1, 'Main_left')
    #make_tab_body(new_wb[org_sn], gen_cr(org), workplace, 'Main_left')
    #make_tab_body(new_wb[org_sn], gen_cr(org), info_base["Operating System:"], 'Main_left')
    #make_tab_body(new_wb[org_sn], gen_cr(org), oper_sys[0], 'Main_center')
    #make_tab_body(new_wb[org_sn], gen_cr(org), info_cpu_2["CPU Brand Name:"], 'Main_left')
    #make_tab_body(new_wb[org_sn], gen_cr(org), info_cpu_1["Number Of Processor Cores:"], 'Main_center')
    #make_tab_body(new_wb[org_sn], gen_cr(org), info_cpu_2["CPU Platform:"], 'Main_center')

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
            if j > 7:
                break
        new_wb.save(dst_path + dst_file)
        if i > 5:
            break


















# ========================================

    new_wb.save(dst_path + dst_file)
    print('Save file')



