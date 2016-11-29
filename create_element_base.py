#!/usr/bin/env python

import csv
from natsort import natsorted
import xlsxwriter
import subprocess

from config import CONFIG


def get_csv_content(file, encoding, delimiter):
    with open(file, encoding=encoding, newline='') as table:
        return [content for content in csv.reader(table, delimiter=delimiter)]


def get_elements(content):
    columns = CONFIG['csv_file']['columns']
    # delete stuff from the list [the first row of the content : the row with column names (not including them)]
    del content[0:content.index(columns)+1]
    return content


def filter_elements(elements):
    for element in elements:
        element.pop(0)
        element.pop(5)
    return elements


def convert2json(elements_list):
    jsoned_elements = {}
    for element in elements_list:
        jsoned_element = {}
        # we don`t need the name of the element
        element_params = element[1:]

        jsoned_element[element[0]] = element_params

        jsoned_elements.update(jsoned_element)

    return jsoned_elements


def organize_to_groups(elements):

    groups = CONFIG['element_groups'].keys()
    elements_grouped = {}

    for group in groups:
        group_of_elements = [{element_name: elements[element_name]} for element_name in elements if group in elements[element_name]]
        elements_grouped.update({group: group_of_elements})

    return elements_grouped


def extract_group(elements, group):
    return dict([(key, d[key]) for d in elements[group] for key in d])


def sort_group_naturally(element_group):
    return natsorted(element_group)


def prepare_group_for_table(elements_group, group_name):
    elements = sort_group_naturally(elements_group)

    prev_name = elements[0]

    ready_names = []
    ready_params = []
    ready_count = []
    ready_manufact = []

    equal_count = 0

    start_name = elements[0]
    last_name = elements[0]

    for cur_name in elements:

        cur_params = elements_group[cur_name]
        prev_params = elements_group[prev_name]

        # elements are the same
        if cur_params == prev_params:
            # and e is first seen
            if equal_count == 0:
                # is it first element? Don`t touch it
                if cur_name == elements[0]:
                    continue
                # is it last element? Add it and stop
                if cur_name == elements[-1]:
                    ready_names.append(cur_name)
                    ready_params.append([cur_params[1], cur_params[3], cur_params[0]])
                    ready_count.append(str(equal_count+1))
                    ready_manufact.append(cur_params[2].split())
                    break
                start_name = prev_name
            # and e is seen before
            elif equal_count > 0:
                # is it last element? So add it and stop
                if cur_name == elements[-1]:
                    ready_names.append('%s..%s' % (start_name, cur_name))
                    ready_params.append([cur_params[1], cur_params[3], cur_params[0]])
                    ready_count.append(str(equal_count+2))
                    ready_manufact.append(cur_params[2].split())
                    break

            last_name = cur_name
            equal_count += 1

        # elements are different
        else:
            # only one
            if equal_count == 0:
                ready_names.append(prev_name)
                ready_params.append([cur_params[1], cur_params[3], cur_params[0]])
                ready_count.append(str(equal_count+1))
                ready_manufact.append(cur_params[2].split())
            # not only one
            else:
                ready_names.append('%s..%s' % (start_name, last_name))
                ready_params.append([cur_params[1], cur_params[3], cur_params[0]])
                ready_count.append(str(equal_count+1))
                ready_manufact.append(cur_params[2].split())

            equal_count = 0
            start_name = ''
            last_name = ''
            prev_name = cur_name

    group_table = []
    for n,p,c,m in zip(ready_names, ready_params, ready_count, ready_manufact):
        group_table.append((n,p,c,m))

    return group_table


def create_xlsx(elements_data, output_file):

    def add_empty_row(col_range):
        global row
        row += 1
        for c in range(col_range):
            worksheet.write(row, c, '', cell_format)

    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    centered_format = workbook.add_format({'border':1})
    centered_format.set_align('center')

    header_format = workbook.add_format({'border':1})
    header_format.set_align('center')
    worksheet.set_row(0, 30, header_format)

    cell_format = workbook.add_format({'border':1})

    global row
    row = 0
    col = 0

    worksheet.set_column(0, 0, 8)
    worksheet.set_column(1, 1, 40)
    worksheet.set_column(2, 2, 6)
    worksheet.set_column(3, 3, 25)

    worksheet.write(row, col, 'Обоз-\nначение', header_format)
    worksheet.write(row, col+1, 'Наименование', header_format)
    worksheet.write(row, col+2, 'Кол.', centered_format)
    worksheet.write(row, col+3, 'Примечание', centered_format)

    add_empty_row(col_range=4)
    add_empty_row(col_range=4)

    for group in elements_data:
        worksheet.write(row, col + 1, CONFIG['element_groups'][group], centered_format)
        row += 1

        for element in elements_data[group]:
            # name
            worksheet.write(row, col, str(element[0]), cell_format)
            # params
            worksheet.write(row, col+1, ','.join(element[1]), cell_format)
            # count
            worksheet.write(row, col+2, str(element[2]), cell_format)
            # manufacturer
            if len(element[3]) > 1:

                for word in range(len(element[3])):
                    worksheet.write(row, col+3, element[3][word], cell_format)

                    if word != range(len(element[3]))[-1]:
                        add_empty_row(col_range=3)
                    else:
                        continue
            else:
                worksheet.write(row, col + 3, ''.join(element[3]), cell_format)

            add_empty_row(col_range=4)

        add_empty_row(col_range=4)

    workbook.close()


def convert_xlsx2pdf(input_file):
    try:
        subprocess.check_output('%s --headless --convert-to pdf %s' %
                        (CONFIG['libreoffice_bin'], input_file),
                         stderr=subprocess.STDOUT, shell=True)
        print('PDF-file created!')
    except subprocess.CalledProcessError as err:
        print('Can`t convert XLSX to PDF using LibreOffice! Check CONFIG.')
        print(err)


def main():
    try:

        content = get_csv_content(CONFIG['csv_file']['filename'],
                                  CONFIG['csv_file']['encoding'],
                                  CONFIG['csv_file']['delimiter'])

        elements = get_elements(content)

        filter_elements(elements)

        jsoned_elements = convert2json(elements)

        organized_groups = organize_to_groups(jsoned_elements)

        group_res = extract_group(organized_groups, 'RES')
        group_cap = extract_group(organized_groups, 'CAP')
        group_ch = extract_group(organized_groups, 'CH')
        group_co = extract_group(organized_groups, 'CO')

        prepared_group_res = prepare_group_for_table(group_res, 'RES')
        prepared_group_cap = prepare_group_for_table(group_cap, 'CAP')
        prepared_group_ch = prepare_group_for_table(group_ch, 'CH')
        prepared_group_co = prepare_group_for_table(group_co, 'CO')

        prepared_elements = {'RES': prepared_group_res,
                             'CAP': prepared_group_cap,
                             'CH' : prepared_group_ch,
                             'CO' : prepared_group_co}

        create_xlsx(prepared_elements, 'elements.xlsx')
        convert_xlsx2pdf('elements.xlsx')

    except ImportError:
        print("Something goes wrong! Check CONFIG.")


if __name__ == '__main__':
    main()
