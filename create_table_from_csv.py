#!/usr/bin/env python

import csv

CONFIG = {
  "elements": 
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


def convert2json(elements):
  
  columns = CONFIG['csv_file']['columns']
  needed_colums = CONFIG['csv_file']['needed_colums']
  
  jsoned_elements = []
  
  for element in elements:
    jsoned_element = {}
    for column, element_parm in zip(columns, element):
      if column in needed_colums:
        jsoned_element[column] = element_parm
    jsoned_elements.append(jsoned_element)
    
  print(jsoned_elements)
  return jsoned_elements


def organize(elements):
  pass
  
  

def main():
  
  try:
    csv_file = CONFIG['csv_file']['filename']
    encoding = CONFIG['csv_file']['encoding']
    delimiter = CONFIG['csv_file']['delimiter']
    
    content = get_csv_content(csv_file, encoding, delimiter)
    
    elements = get_elements(content)
    # print(elements)
    print(filter_elements(elements))
    #convert2json(elements)
    
  except ImportError:
    print("Something goes wrong with parsing %s" % CONFIG['csv_file']['filename'])


if __name__ == '__main__':
  main()
