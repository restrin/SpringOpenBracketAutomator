# -*- coding: utf-8 -*-
"""
Created on Sat Apr 18 17:29:48 2015

@author: Ron
"""

import shutil
import copy
from openpyxl import load_workbook

# Assume that bracket numbers are ordered as a binary tree
#       1
#     2  3
#   4 5 6 7
# ...

# Maps the fighter number to their position in the bracket
fighter_to_excel_map = ['L19', # Winner
                        'J11', 'J27', #Finals
                        'H8', 'H14', 'H24', 'H30', #Semifinals
                        'F6', 'F10', 'F12', 'F16', 'F22', 'F26', 'F28', 'F32'];                                    

def get_bracket_indices(n):
    '''
        Returns indices in the bracket for placing n competitors
    '''
    return range(n,2*n)

def get_competitors_per_division(ws, start):
    '''
        Reads worksheet ws to obtain the number of competitors in
        the division starting from row 'start'. Divisions are
        separated by empty rows.
    '''
    count = 0
    while(ws['A'+str(start+count)].value != None):
        count += 1
    return count

def get_competitors(sparws, start, count):
    '''
        Reads rows 'start' to 'start+count' to obtain the names
        and schools of competitors in the division defined by those
        rows. Result is stored in the list 'competitors'
    '''
    competitors = []
    for i in range(count):
        first_name = sparws['B' + str(start+i)].value
        last_name = sparws['A' + str(start+i)].value
        school = sparws['P' + str(start+i)].value
        if (school != None):
            competitors.append((first_name + ' ' + last_name, school))
        else:
            competitors.append((first_name + ' ' + last_name, ''))
    return competitors

def write_competitors_to_bracket(ws, competitors, indices, outwb, output_fname):
    for i in range(len(competitors)):
        c = competitors[i]
        text = c[0] + ', ' + c[1]
        ws[fighter_to_excel_map[indices[i]-1]] = text
    outwb.save(output_fname)

def fill_in_brackets(template_fname, data_fname, output_fname, copyFileFlag):
    if (copyFileFlag):
        # Make copy of Bracket Template
        shutil.copyfile(template_fname, output_fname)
    
    sparwb = load_workbook(data_fname) # Bracket data
    sparws = sparwb.active
    outwb = load_workbook(output_fname) # Output file
    
    ix = 2;
    division = 0;
    num_per_division = get_competitors_per_division(sparws, ix)
    
    while (num_per_division > 0):
        competitors = get_competitors(sparws, ix, num_per_division)
        bracket_indices = get_bracket_indices(num_per_division)
        
        ix += num_per_division + 1
        num_per_division = get_competitors_per_division(sparws, ix)
        
        outws = outwb.worksheets[division]        
        outws.title = 'Division ' + str(division)        
        
        if (num_per_division > 0):
            # Only make a new template copy if there's another division
            outwb.add_sheet(copy.deepcopy(outws), division+1)

        write_competitors_to_bracket(outws, competitors, bracket_indices, outwb, output_fname)
        division += 1
        
        print division

    outwb.save(output_fname)