# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead
"""

import comtypes.client as CC

def open_mathcad():
    ccHandle = CC.CreateObject("MathcadPrime.Application")
    print (ccHandle)
    ccHandle.Worksheet

open_mathcad()
