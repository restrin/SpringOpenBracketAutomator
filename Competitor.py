# -*- coding: utf-8 -*-
"""
Created on Sun Apr 19 15:07:46 2015

@author: Ron
"""

class Competitor(object):
    def __init__(self):
        self.first_name = ''
        self.last_name = ''
        self.school = ''
        self.gender = ''
        self.age = 0
        self.weight = 0
        
class SparringCompetitor(Competitor):    
    Belts = ['yellow', 'green', 'blue', 'red', 'black']
    def __init__(self):
        self.belt = None