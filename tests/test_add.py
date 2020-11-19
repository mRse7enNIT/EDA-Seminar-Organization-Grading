# -*- coding: utf-8 -*-
# @Time    : 11/19/20 7:40 PM
# @Author  : Saptarshi
# @Email   : saptarshi.mitra@tum.de
# @File    : test_add.py
# @Project: eda-seminar-organization-grading

from toolscripts.ConfirmedStudents import add

print("before test")
print(add(15,16))

#function for testing in pytest
def test_add():
    res = add(1, 1)
    assert res == 2
    print("after test")
