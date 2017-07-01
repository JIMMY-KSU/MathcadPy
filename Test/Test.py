# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead
"""
import comtypes.client as CC
import win32com.client as win32
import pythoncom
import os

def open_mathcad():
    print(win32.gencache.EnsureModule("MathcadPrime.Application", 0, 1, 2))
    mcad = win32.Dispatch("MathcadPrime.Application")
    mcad.Visible = True

#open_mathcad()

#print("\n\n")

def open_mathcad2():
    ccHandle = CC.CreateObject("Ptc_MathcadPrime_Automation")
    print(ccHandle)
    ccHandle.Worksheet

#open_mathcad2()

def register_mathcad_com():
    methods = {}
    try:
        tlbpath = r"C:\Program Files\PTC\Mathcad Prime 3.1\Ptc.MathcadPrime.Automation.tlb"
        mcad = pythoncom.LoadTypeLib(tlbpath)
        pythoncom.RegisterTypeLib(mcad, tlbpath)
        for i in range(mcad.GetTypeInfoCount()):
            obj = mcad.GetDocumentation(i)[0]  # COM object name
            CLSID = mcad.GetTypeInfo(i).GetTypeAttr().iid.__str__()  # CLSID
            methods[obj] = CLSID
        print("\nRegistered:")
        for k, v in methods.items():
            if pythoncom.IsGatewayRegistered(v):
                print(k)
        return methods
    except pythoncom.com_error:
        print("COM error")
        return None

register_mathcad_com()

def attempt_cast_coclass():
    mcad = win32.Dispatch("MathcadPrime.Application")
    mcad.Visible = True
    mcad.Open(os.path.join(os.getcwd(), "test.mcdx"))
    Outputs = win32com.client.CastTo(mcad, "Outputs")
    print(Outputs.Count)

attempt_cast_coclass()