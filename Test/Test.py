# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead
"""

import win32com.client as win32

def open_mathcad():
    mcad = win32.gencache.EnsureModule("MathcadPrime.Application", 0, 1, 2)
    mcad = win32.Dispatch("MathcadPrime.Application")
    mcad.Visible = True


def open_mathcad2():
    #mathcad = win32.gencache.EnsureDispatch("Ptc_MathcadPrime_Automation.Application")
    import comtypes.client as CC

    ccHandle = CC.CreateObject("MathcadPrime.Application")
    print (ccHandle)
    ccHandle.Worksheet


#open_word()
open_mathcad()
print("\n\n")
open_mathcad2()