#!/usr/bin/env python

import csv
from collections import OrderedDict
from natsort import natsorted


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
    return dict([(key,d[key]) for d in elements[group] for key in d])


def sort_group_naturally(element_group):
    return natsorted(element_group)


def create_table(element_group):
    start_element = sort_group_naturally(element_group)[0]
    print('start %s' % start_element)
    same_count = 0
    for element_name in sort_group_naturally(element_group):
        if element_group[element_name] == element_group[start_element]:
            if same_count > 0:
                pass
            else:
                print('..')
            same_count += 1
        else:
            if same_count == 0:
                print('only one %s' % start_element)
                print('\nstart %s' % element_name)
            else:
                print('end %s' % start_element)
                print('\nstart %s' % element_name)
                same_count = 0
        start_element = element_name
    print(start_element)
#        print(element_name, element_group[element_name])


def main():
    try:

        content = get_csv_content(CONFIG['csv_file']['filename'],
                                  CONFIG['csv_file']['encoding'],
                                  CONFIG['csv_file']['delimiter'])

        elements = get_elements(content)

        filter_elements(elements)

        jsoned_elements = convert2json(elements)

        organized_groups = organize_to_groups(jsoned_elements)

        elements_res = extract_group(organized_groups,'RES')
        print(elements_res)
#        elements_cap = extract_group('CAP')
#        elements_ch = extract_group('CH')
#        elements_co = extract_group('CO')

       # create_table(elements_res)




    except ImportError:
        print("Something goes wrong with parsing %s" % CONFIG['csv_file']['filename'])


if __name__ == '__main__':
    main()
