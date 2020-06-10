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
    panel_ziyan = ['小', '26', '26基因', '26gene', '26.0', '结直肠癌12基因']
    panel_jiezhichangai = ['结直肠癌14个驱动基因', '结直肠癌14驱动基因',
                           '结直肠癌NCCN指南', '结直肠癌NCCN指南必选检测',
                           '结直肠癌14个驱动基因检测']
    panel_62gene = ['中']
    panel_4gene = ['4基因', '4gene']
    panel_ckit = ['ckit', 'c-kit', 'C-KIT', 'CKIT', 'C-KIT析', 'CIKT']
    panel_rna = ['超级全外RNA融合']
    panel_ignore = ['肿瘤易感', '城市易感', '软组织肉瘤基因检测',
                    'TMB', 'BRCA1/2', '遗传', '淋巴瘤48基因']

    if panel in panel_ziyan:
        return 'ziyan'
    elif panel in panel_jiezhichangai:
        return 'jiezhichangai'
    elif panel in panel_62gene:
        return '62gene'
    elif panel in panel_4gene:
        return '4gene'
    elif panel in panel_ckit:
        return 'ckit'
    elif panel in panel_rna:
        return ''
    elif panel in panel_ignore:
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
        output_txt.append([int(i['序号']), str(i['检测编号']).upper(), i['DNA标签'], i['RNA标签'], get_panel(str(i['检测项目']))])
    write_file(options.output, output_txt)
