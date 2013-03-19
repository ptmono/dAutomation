#!/usr/bin/python
# coding: utf-8

from os.path import dirname, abspath

PICKITEMSP = True
RAREP	   = True
REPAIRP    = False

ITEMS = {
    "legendary": ["#02CE01",    # set
                  "#BF642F"],   # legndary
    "rare":	 ["#BBBB00"]
    }

current_abpath = abspath(dirname(__file__)) + "/"
# With py2exe the dirname is INSTPATH/server/library.zip. So
# current_abpath will be INSTPATH/server/library.zip/
if current_abpath[-12:] == "library.zip/":
    current_abpath = current_abpath[:-12]

imgs_dir = current_abpath + "imgs\\"


def get_item_colors():
    '''
    >>> get_item_colors()
    '''
    result = []
    if not PICKITEMSP: return result
    
    if RAREP:
        for a in ITEMS:
            result += ITEMS[a]
        return result
    else:
        result = ITEMS["legendary"]
        return result
        
