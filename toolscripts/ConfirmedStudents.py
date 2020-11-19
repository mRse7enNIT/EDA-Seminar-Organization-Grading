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


def read_srcfile(source_filename):
    # Import the src file from TUMonline into a Pandas dataframe
    src_xlsx = pd.ExcelFile(source_filename)
    print("Sheets present in the src xlsx file: \n")
    print(src_xlsx.sheet_names)
    src_df = src_xlsx.parse('TUMonline_2020-11-07')
    return src_df


def main():
    source_filename = sys.argv[1]
    src_df = read_srcfile(source_filename)
    print(src_df.to_string())

if __name__ == "__main__":
    main()
