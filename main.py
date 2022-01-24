# -*- coding: utf-8 -*-
"""
Created on Fri Jan 21 23:49:50 2022

@author: dejong71
"""

# ------------------------------------------------------------------------------

import os
import pandas as pd
from dataclasses import dataclass

# ------------------------------------------------------------------------------

def floor(x: int, div: int):
    '''Method to floor floor an integer (x) by the divisor (div).'''
    x -= x % div
    return x

# ------------------------------------------------------------------------------

def verify_number(number : str):
    split = number.split('+')[-1]
    try:
        temp = int(split[2:])
        return True if temp >= 1e7 else False
    except ValueError:
        return False

# ------------------------------------------------------------------------------

def format_number(number : str):
    '''Method to filter out numbers shorter than 8 digits.'''
    split = number.split('+')[-1]
    try:
        temp = int(split[2:])
        return temp
    except ValueError:
        return split
        
# ------------------------------------------------------------------------------

@dataclass
class PhoneNumber:
    '''A dataclass to model a phonenumber.\n
    - number : int : The formatted phonenumber.\n
    - correct : bool : Flag correctly (True) and incorrectly (False) formatted numbers.\n
    - found : bool : Flag if the number has been found in a database (True) or not (False).\n
    - table : str : The name of the sheet where the number was found.\n
    - group : int : The size of the group to whicht the number belongs.\n
    - flag : bool : The status of the VMS flag of the number group.'''
    number: int
    correct: bool
    found: bool
    table: str
    group: int
    flag: bool

# ------------------------------------------------------------------------------

# Locate the datafolder.
folder = os.getcwd()
# Construct a master database containing both sheets.
path_dbase = os.path.join(folder, 'tg.xlsx')
sheets = ('tg1', 'tg2')
frames = []
for sheet in sheets:
    df = pd.read_excel(path_dbase, sheet)
    df['Sheet'] = sheet
    frames.append(df)
dbase_in = pd.concat(frames, ignore_index=True)
# Clean the input numbers.
path_numbers = os.path.join(folder, 'I-Mens-Gzat.xlsx')
numbers = pd.read_excel(path_numbers)
numbers['Correct'] = numbers['Access number'].apply(verify_number)
numbers['Formatted'] = numbers['Access number'].apply(format_number)
# Keep track of previously found numbers.
found = set()
# Save numbers in three categories.
numbers_found = {}
numbers_missing = {}
numbers_wrong = {}

dbase = dbase_in.copy()

# Loop over all available group sizes of numbers from large to small.
for size in dbase.Size.unique()[::-1]:    
    # Loop over all correctly formatted numbers that have not been found yet.
    for number in numbers.Formatted[numbers.Correct == True]:
        # If the number has not yet been found.
        if number not in found:
            # Floor the number, unless floored by 1.
            floored = floor(number, size) if size > 1 else number
            # If the floored number is found in the size database.
            if floored in set(dbase.Start[dbase.Size == size]):
                # Track the flag and table of origin.
                flag = dbase.VMS[dbase.Start == floored].item()
                table = dbase.Sheet[dbase.Start == floored].item()
                # Add to the list of found numbers.
                numbers_found[floored] = PhoneNumber(
                    floored, True, True, table, size, flag)
                found.add(number)
            # If the floored number is not found in the size database.
            else:
                # Register the number as unfound if left in the last loop.
                if size == 1:
                    numbers_missing[number] = PhoneNumber(
                        number, True, False, None, None, None)
    # Incorrectly formatted numbers.
    for number in numbers.Formatted[numbers.Correct == False]:
        numbers_wrong[number] = PhoneNumber(
            number, False, None, None, None, None)
# Save as csv.
with open('output.csv', 'w') as f:
    f.write('number,correct,found,table,group,vms\n')
    for numberset in numbers_found, numbers_missing, numbers_wrong:
        for number in numberset.values():
            f.write(
                f'{number.number},{number.correct},{number.found},{number.table},{number.group},{number.flag}\n')