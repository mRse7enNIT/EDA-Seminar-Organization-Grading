# -*- coding: utf-8 -*-
# @Time    : 11/19/20 7:28 PM
# @Author  : Saptarshi
# @Email   : saptarshi.mitra@tum.de
# @File    : ConfirmedStudents.py
# @Project: eda-seminar-organization-grading

#def add(a: int, b: int) -> int:
#   return a + b

import numpy as np
import pandas as pd
import sys
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import *
import sys

def read_srcfile(source_filename):
    # Import the src file from TUMonline into a Pandas dataframe
    #src_xlsx = pd.ExcelFile(source_filename)
    #print("Sheets present in the src xlsx file: \n")
    #print(src_xlsx.sheet_names)
    #src_df = src_xlsx.parse('TUMonline_2020-11-07')
    src_wb = load_workbook(source_filename)
    print("The available sheets in the source xlsx file")
    print(src_wb.sheetnames)
    src_sheet = src_wb.active
    print("selected sheet for data manipulation:")
    print(src_sheet)
    src_df = pd.DataFrame(src_sheet.values)
    return src_df


def write_masterfile(write_df):
    # master_wb = Workbook()
    # current_ws = master_wb.active
    # for r in dataframe_to_rows(write_df, index=True, header= True):
    #     current_ws.append(r)
    # for cell in current_ws['A'] + current_ws[1]:
    #     cell.style = 'Pandas'
    #
    # master_wb.save("OutputFiles/master_sheet_" + str(pd.datetime.now().strftime('%Y_%m_%d_%H_%M_%S')) + ".xlsx")
    writer = pd.ExcelWriter("OutputFiles/master_sheet_" + str(pd.datetime.now().strftime('%Y_%m_%d_%H_%M_%S')) + ".xlsx", engine='xlsxwriter')
    write_df.to_excel(writer,'Sheet1')
    writer.save()


def replace_header(input_df):
    new_header = input_df.iloc[0]
    input_df = input_df[1:]
    input_df.columns=new_header
    return input_df


def choose_fixed_place(modified_df):
    # for (idx,row) in modified_df.iterrows():
    #     print(idx,row)
    filtered_df = modified_df[(modified_df.STATUS=='Fixplatz')]
    filtered_df = filtered_df[['FAMILIENNAME','VORNAME','MATRIKELNUMMER','GESCHLECHT','E-MAIL']]
    return filtered_df


def add_columns(filtered_df):
    print("adding additional columns for TITEL, BETREUER, VORTRAG")
    extended_df = filtered_df.assign(TITEL='', BETREUER='', VORTRAG='')
    return extended_df


def add_columns_HS(extended_df):
    print("adding additional columns for REVIEW FÜR (AKTIVE), REVIEW VON (PASSIV)")
    extended_df = extended_df.assign(REVIEW_FÜR='', REVIEW_VON='')
    return extended_df


def main():
    arguments = len(sys.argv) - 1
    position = 1
    while (arguments >= position):
        print("Parameter %i: %s" % (position, sys.argv[position]))
        position = position + 1
    source_filename = sys.argv[1]
    src_df = read_srcfile(source_filename)
    print(src_df.to_string())
    modified_df = replace_header(src_df)
    print(modified_df)
    filtered_df =choose_fixed_place(modified_df)
    print(filtered_df)
    extended_df = add_columns(filtered_df)
    if(arguments==2 and sys.argv[2]=='-HS'):
        extended_df = add_columns_HS(extended_df)
    write_masterfile(extended_df)


if __name__ == "__main__":
    main()
