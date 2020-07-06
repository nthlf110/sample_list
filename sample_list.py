import os
import sys
import xlwt
import xlrd
import optparse
import csv


def read_from_xlsx(xls_file_name, sheet_name_in_xls, header="T"):
    data_info = xlrd.open_workbook(xls_file_name)
    try:
        data_sh = data_info.sheet_by_name(sheet_name_in_xls)
    except:
        print("no sheet in %s named data" % xls_file_name)

    data_nrows = data_sh.nrows
    data_ncols = data_sh.ncols

    result_data = []
    if header == "T":
        for i in range(1, data_nrows):
            result_data.append(dict(zip(data_sh.row_values(0), data_sh.row_values(i))))
    elif header == "F":
        for i in range(data_nrows):
            result_data.append(data_sh.row_values(i))
    else:
        print("header = ", header)
        print("The parameter 'header' is undefined, please check!")
        os._exit()

    return result_data


def write_file(path, data):
    try:
        with open(path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, delimiter='\t')
            for row in data:
                writer.writerow(row)
    except Exception as e:
        print("Write to: %s, ERROR: %s" % (path, e))


def get_panel(panel):
    panel_4gene = ['4基因', '4gene']
    panel_ignore = ['易感', '儿童', '肉瘤',
                    'TMB', 'BRCA', '遗传', '淋巴',
                    '前列腺']

    if panel.__contains__('26') or panel == '小' or panel.__contains__('12'):
        return 'ziyan'
    elif panel.__contains__('14') or panel.__contains__('NCCN'):
        return 'jiezhichangai'
    elif panel == '中':
        return '62gene'
    elif panel in panel_4gene:
        return '4gene'
    elif sum([panel.upper().__contains__(i) for i in ['C','I','K','T']]) == 4:
        return 'ckit'
    elif panel.__contains__('RNA'):
        return ''
    elif sum([panel.__contains__(i) for i in panel_ignore]):
        return 'ignore'
    else:
        return 'unknown'


if __name__ == '__main__':
    parser = optparse.OptionParser()
    parser.add_option('-i', '--input', dest='input')
    parser.add_option('-o', '--output', dest='output', default='./sample.list.txt')
    (options, args) = parser.parse_args()
    input_xlsx = read_from_xlsx(options.input, 'Sheet1', header='T')
    output_txt = []
    for i in iter(input_xlsx):
        if isinstance(i['DNA标签'], float):
            i['DNA标签'] = int(i['DNA标签'] )
        if isinstance(i['RNA标签'], float):
            i['RNA标签'] = int(i['RNA标签'] )
        if isinstance(i['DNA标签'], str):
        	i['DNA标签'] = '/'
        if isinstance(i['RNA标签'], str):
            i['RNA标签'] = '/'
        output_txt.append([int(i['序号']), str(i['检测编号']).upper(), i['DNA标签'], i['RNA标签'], get_panel(str(i['检测项目']))])
    write_file(options.output, output_txt)
