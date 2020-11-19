# -*- coding: utf-8 -*-

from setuptools import setup, find_packages


with open('README.md') as f:
    readme = f.read()

with open('LICENSE') as f:
    license = f.read()

setup(
    name='EDA-Seminar-Organization-Grading',
    version='0.1.0',
    description='Project comprising scripts for management of Seminars at the EDA chair',
    long_description=readme,
    author='Saptarshi Mitra',
    author_email='saptarshi.mitra@tum.de',
    url='https://gitlab.lrz.de/ga53wis/eda-seminar-organization-grading',
    license=license,
    packages=find_packages(exclude=('tests', 'docs'))
)

