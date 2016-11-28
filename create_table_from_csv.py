#!/usr/bin/env python

import csv
from natsort import natsorted
import xlsxwriter


CONFIG = {
    "element_groups":
        {"CAP": "Конденсаторы",
         "RES": "Резисторы",
         "CH": "Микросхемы",
         "CO": "Индуктивности"},

    "csv_file":
        {"filename": "dm66_keyboard.csv",
         "encoding": "cp1251",
         "delimiter": ";",
         "columns": ['Count', 'RefDes', 'Value', 'PartNumber', 'Manufacturer', 'Package', 'ROHS', 'N_GROUP'],
         "needed_colums": ['RefDes', 'Value', 'PartNumber', 'Manufacturer', 'Package', 'N_GROUP']}}


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


def create_table(elements_data):
    workbook = xlsxwriter.Workbook('elements.xlsx')
    worksheet = workbook.add_worksheet()
    format = workbook.add_format()
    format.set_align('center')
    content_format = workbook.add_format()
    worksheet.set_row(0, 30, content_format)

    row = 0
    col = 0

    worksheet.set_column(0, 0, 8)
    worksheet.set_column(1, 1, 40)
    worksheet.set_column(2, 2, 6)
    worksheet.set_column(3, 3, 25)

    worksheet.write(row, col, 'Обоз-\nначение', content_format)
    worksheet.write(row, col+1, 'Наименование', format)
    worksheet.write(row, col+2, 'Кол.', format)
    worksheet.write(row, col+3, 'Примечание', format)

    row += 2

    for group in elements_data:
        worksheet.write(row, col + 1, CONFIG['element_groups'][group], format)
        row += 1

        for element in elements_data[group]:
            # name
            worksheet.write(row, col, str(element[0]))
            # params
            worksheet.write(row, col+1, ','.join(element[1]))
            # count
            worksheet.write(row, col+2, str(element[2]))
            # manufacturer
            if len(element[3]) > 1:
                for word in range(len(element[3])):
                    worksheet.write(row, col+3, element[3][word])
                    if word != range(len(element[3]))[-1]:
                        row += 1
                    else:
                        continue
            else:
                worksheet.write(row, col + 3, ''.join(element[3]))
            row += 1

        row += 1

    workbook.close()


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

        create_table(prepared_elements)

    except ImportError:
        print("Something goes wrong with parsing %s" % CONFIG['csv_file']['filename'])


if __name__ == '__main__':
    main()
