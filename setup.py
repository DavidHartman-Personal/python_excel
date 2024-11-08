# -*- coding: utf-8 -*-


from setuptools import setup, find_packages


with open('README.rst') as f:
    readme = f.read()

with open('LICENSE') as f:
    license = f.read()

setup(
    name='excel_workbook',
    version='0.1.0',
    description='Python excel utilities',
    long_description=readme,
    author='David Hartman',
    author_email='david.j.hartman@gmail.com',
    url='TBD',
    license=license,
    packages=find_packages(exclude=('tests', 'docs'))
)

