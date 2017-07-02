# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead

Requirements:

Mathcad Prime
Python Win 32 extensions (https://sourceforge.net/projects/pywin32/)

"""

import comtypes.client as CC
import comtypes

"""
@TODO

In the registry, "MathcadPrime.Application" is linked to the
Ptc.MathcadPrime.Automation.Application part of the TLB file. Need to find a
way to access the other top level functions that are part of
Ptc.MathcadPrime.Automation (e.g. worksheet, inputs, outputs)

Currently considering creating a new registry key, or finding a win32 or
Comtypes method that allows stepping back a level in the com library.
"""

# First run - register
# @TODO add something to decide if it is already registered
#    try:
#        tlbpath = r"C:\Program Files\PTC\Mathcad Prime 3.1\Ptc.MathcadPrime.Automation.tlb"
#        mcad = pythoncom.LoadTypeLib(tlbpath)
#        pythoncom.RegisterTypeLib(mcad, tlbpath)
#    except pythoncom.com_error:
#        print("Encountered COM error when registering Mathcad COM functions")


class Mathcad(object):
    """ Top level Mathcad wapper class """
    def __init__(self, visible=False):
        try:
            self.__mcadapp = CC.CreateObject("MathcadPrime.Application")
            self.version = self.__mcadapp.GetVersion()  # Fetches Mathcad version
            if visible is False:
                self.__mcadapp.Visible = False
            else:
                self.__mcadapp.Visible = True
        except:
            print("Could not locate the Mathcad Automation API")

    def activate(self):
        """ Activate the Mathcad window. If visible, this maximises Mathcad"""
        self.__mcadapp.Activate()

    def active_sheet(self):
        """ Returns the active worksheet name """
        name = self.__mcadapp.ActiveWorksheet.Name
        if name == "":
            return None  # Returns none if the current worksheet not saved
        else:
            return name

    def worksheets(self):
        """ lists worksheets open in the Mathcad instance """
        worksheets = []
        for i in range(self.__mcadapp.Worksheets.Count):
            print(self.__mcadapp.Worksheets.Item(i).Name)
            worksheets.append(self.__mcadapp.Worksheets.Item(i).Name)
        return worksheets

    def close_all(self, save_option="Discard"):
        """ Closes all worksheets. Can specify save options before closing """
        if save_option in ["Discard", 2]:
            self.__mcadapp.CloseAll(2)
        elif save_option in ["Prompt", 1]:
            self.__mcadapp.CloseAll(1)
        elif save_option in ["Save", 0]:
            self.__mcadapp.CloseAll(0)
        else:
            print("incorrect save argument specified")

#class Worksheet(object):
#    """ Mathcad Worksheet class """
#    def __init__(self, filepath=None):
#        self.__mcadapp = CC.CreateObject("MathcadPrime.Application")
#        self.__worksheet = comtypes.cast(self.__mcadapp)
#        if filepath is not None:
#            try:
#                self.__mcadapp.Open(filepath)
#            except:
#                print("error opening file")
#        self.name = self.__worksheet.Name
#
#    def activate(self):
#        """ activates the worksheet object """
#        self.__worksheet.Activate()
#
#    def IsOpen(self):
#        """ Opens a worksheet file """
#        return self.__worksheet.IsOpen()
#
#    def Close(self, save_option="Discard"):
#        """ Closes the worksheet """
#        if save_option in ["Discard", 2]:
#            self.__worksheet.CloseAll(2)
#        elif save_option in ["Prompt", 1]:
#            self.__worksheet.CloseAll(1)
#        elif save_option in ["Save", 0]:
#            self.__worksheet.CloseAll(0)
#        else:
#            print("incorrect save argument specified")
#
#    def set_real_input(self, input_alias, value, units):
#        """ Set the value of a numerical input range in the worksheet """
#        if input_alias in self.worksheet.Inputs:
#            self.worksheet.SetRealValue(str(input_alias), value, str(units))


if __name__ == "__main__":

    import os
    test = os.path.join(os.getcwd(), "Test", "test.mcdx")

    print("Example usage")
    MC = Mathcad(visible=False) # Open Mathcad with no GUI
    print(MC.active_sheet())  # Print the name of the active sheet
    MC.activate()
    print(MC.version)
    print(MC.worksheets())
