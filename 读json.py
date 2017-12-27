# -*- coding: utf-8 -*-
"""
Created on Thu Dec 21 15:26:13 2017

@author: user
"""

import json
f1= open(r"json1.json",encoding='utf-8')
new_f1 = json.load(f1)
print(new_f1)
f2 = open(r"json2.json",encoding='utf-8')
new_f2 = json.load(f2)
print(new_f2)
f3 = open(r"json3.json",encoding='utf-8')
new_f3 = json.load(f3)
print(new_f3)


    
