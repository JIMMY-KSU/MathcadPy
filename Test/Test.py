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
    import comtypes

    ccHandle = CC.CreateObject("Ptc_MathcadPrime_Automation.Application")
    ccHandle = CC.CreateObject("Mathcad.Application")
    print (ccHandle)
    import comtypes.gen.CSharpServer as CS
    InterfaceHandle = ccHandle.QueryInterface(CS.IManagedInterface)

    print ("output of PrintHi function = ", InterfaceHandle.PrintHi("World"))

#open_word()

open_mathcad()