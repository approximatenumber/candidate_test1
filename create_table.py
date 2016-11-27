#!/usr/bin/env python

import xlsxwriter
from natsort import natsorted

res_list = {'R35': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R56': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R50': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R8': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R21': ['1K3', 'ERJ2RKF1301X', 'Panasonic', '0402', 'RES'], 'R16': ['1K8', 'ERJ2RKF1801X', 'Panasonic', '0402', 'RES'], 'R42': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R7': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R54': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R45': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R31': ['10K', 'ERJ2RKF1002X', 'Panasonic', '0402', 'RES'], 'R40': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R12': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R48': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R15': ['1K8', 'ERJ2RKF1801X', 'Panasonic', '0402', 'RES'], 'R44': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R32': ['33R', 'ERJ2RKF33R0X', 'Panasonic', '0402', 'RES'], 'R9': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R25': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R46': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R47': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R33': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R28': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R17': ['60R4', 'ERJ2RKF60R4X', 'Panasonic', '0402', 'RES'], 'R3': ['1K3', 'ERJ2RKF1301X', 'Panasonic', '0402', 'RES'], 'R43': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R30': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R24': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R29': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R20': ['1K3', 'ERJ2RKF1301X', 'Panasonic', '0402', 'RES'], 'R51': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R5': ['10K', 'ERJ2RKF1002X', 'Panasonic', '0402', 'RES'], 'R1': ['1K3', 'ERJ2RKF1301X', 'Panasonic', '0402', 'RES'], 'R4': ['1K3', 'ERJ2RKF1301X', 'Panasonic', '0402', 'RES'], 'R37': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R14': ['1K8', 'ERJ2RKF1801X', 'Panasonic', '0402', 'RES'], 'R39': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R27': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R6': ['33R', 'ERJ2RKF33R0X', 'Panasonic', '0402', 'RES'], 'R26': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R23': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R59': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R49': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R36': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R34': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R22': ['1K3', 'ERJ2RKF1301X', 'Panasonic', '0402', 'RES'], 'R55': ['10K', 'ERJ2RKF1002X', 'Panasonic', '0402', 'RES'], 'R53': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R38': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R2': ['1K3', 'ERJ2RKF1301X', 'Panasonic', '0402', 'RES'], 'R52': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R19': ['1K3', 'ERJ2RKF1301X', 'Panasonic', '0402', 'RES'], 'R10': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R18': ['60R4', 'ERJ2RKF60R4X', 'Panasonic', '0402', 'RES'], 'R57': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R41': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R13': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES'], 'R58': ['1K3', 'ERJ3EKF1301V', 'Panasonic', '0603', 'RES'], 'R11': ['2K2', 'ERJ2RKF2201X', 'Panasonic', '0402', 'RES']}


def sort_group_naturally(element_group):
    return natsorted(element_group)


workbook = xlsxwriter.Workbook('elements.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

worksheet.write(row, col + 1, 'Резисторы')
row += 1

elements = sort_group_naturally(res_list)

prev_e = elements[0]
worksheet.write(row, col, prev_e)
row += 1
equal_count = 0

for element in elements:

    print(element)

    e_refdes = element
    e_params = res_list[element]
    e_params_prev = res_list[prev_e]

#    e_partnumber = e_params[1]
#    e_package = e_params[3]
#    e_value = e_params[0]
#    e_count = 1
#    e_manufact = e_params[2]

    refdes_to_write = ""

    if e_params == e_params_prev:

        if element == elements[-1]:

            if equal_count == 0:
                worksheet.write(row, col, e_refdes)
            else:
                worksheet.write(row, col, ".." + e_refdes)

        equal_count += 1

    else:

        if equal_count == 0:
            refdes_to_write += e_refdes
            worksheet.write(row, col, e_refdes)
            row += 1

        else:
            worksheet.write(row, col, ".."+prev_e)
            row += 1
            worksheet.write(row, col, e_refdes)
            equal_count = 0
            row += 1

    prev_e = e_refdes




   # worksheet.write(row, col, e_refdes)
  #  worksheet.write(row, col + 1, "%s, %s, %s" % (e_partnumber, e_package, e_value))
  #  worksheet.write(row, col + 2, e_count)
  #  worksheet.write(row, col + 3, e_manufact)

   # row += 1

workbook.close()
