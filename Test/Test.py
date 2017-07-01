# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead
"""

import win32com.client as win32
import pythoncom
import os

def open_mathcad():
    print(win32.gencache.EnsureModule("MathcadPrime.Application", 0, 1, 2))
    mcad = win32.Dispatch("MathcadPrime.Application")
    mcad.Visible = True


def open_mathcad2():
    #mathcad = win32.gencache.EnsureDispatch("Ptc_MathcadPrime_Automation")
    import comtypes.client as CC

    ccHandle = CC.CreateObject("Ptc_MathcadPrime_Automation")
    print (ccHandle)
    ccHandle.Worksheet


#open_word()
#open_mathcad()
#print("\n\n")
open_mathcad2()

#def register_mathcad_com():
#    ret = []
#    try:
#        tlbpath = r"C:\Program Files\PTC\Mathcad Prime 3.1\Ptc.MathcadPrime.Automation.tlb"
#        mcad = pythoncom.LoadTypeLib(tlbpath)
#        pythoncom.RegisterTypeLib(mcad, tlbpath)
#        for i in range(mcad.GetTypeInfoCount()):
#            obj = mcad.GetDocumentation(i)[0]  # COM object name
#            CLSID = mcad.GetTypeInfo(i).GetTypeAttr().iid.__str__()  # CLSID
#            factory = pythoncom.MakePyFactory(CLSID)
#            regId = pythoncom.CoRegisterClassObject(CLSID, factory, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.REGCLS_MULTIPLEUSE|pythoncom.REGCLS_SUSPENDED)
#            ret.append((factory, regId))
#        return ret
#    except pythoncom.com_error:
#        print("COM error")
#        return None
#
#mcad_methods = register_mathcad_com()
#
#print(mcad_methods)
#
for k, v in mcad_methods:
    if pythoncom.IsGatewayRegistered(v):
        print(k)