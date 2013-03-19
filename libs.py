#!/usr/bin/python
# coding: utf-8

import time

def hex_to_rgb(value):
    value = value.lstrip('#')
    lv = len(value)
    return tuple(int(value[i:i+lv/3], 16) for i in range(0, lv, lv/3))

def rgb_to_hex(rgb):
    return '#%02x%02x%02x' % rgb

# hex_to_rgb("#ffffff")             #==> (255, 255, 255)
# hex_to_rgb("#ffffffffffff")       #==> (65535, 65535, 65535)
# rgb_to_hex((255, 255, 255))       #==> '#ffffff'
# rgb_to_hex((65535, 65535, 65535)) #==> '#ffffffffffff'

def rSleep(t):
    t = t * 0.001
    time.sleep(t)


def sleep(t):
    rSleep(t)
