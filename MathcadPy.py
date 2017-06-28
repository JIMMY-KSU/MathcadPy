# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead

Requirements:

Mathcad Prime
Python Win 32 extensions (https://sourceforge.net/projects/pywin32/)

"""

import win32com.client as win32

"""
@TODO

In the registry, "MathcadPrime.Application" is linked to the
Ptc.MathcadPrime.Automation.Application part of the TLB file. Need to find a
way to access the other top level functions that are part of
Ptc.MathcadPrime.Automation (e.g. worksheet, inputs, outputs)

Currently considering creating a new registry key, or finding a win32 or
Comtypes method that allows stepping back a level in the com library.
"""


class Mathcad(object):
    """ Top level Mathcad wapper class """
    def __init__(self, visible=False):
        try:
            win32.gencache.EnsureModule("MathcadPrime.Application", 0, 1, 2)
            self.mcadapp = win32.Dispatch("MathcadPrime.Application")
            self.version = self.mcadapp.GetVersion()  # Fetches Mathcad version
            if visible is False:
                self.mcadapp.Visible = False
            else:
                self.mcadapp.Visible = True
        except:
            print("Could not locate the Mathcad Automation API")

    class Worksheet(object):
        """ Mathcad Worksheet class """
        def __init__(self, filepath=None):
            win32.gencache.EnsureModule("MathcadPrime.Application", 0, 1, 2)
            self.worksheet = win32.Dispatch("MathcadPrime.Application").Worksheet
            self.inputs = self.worksheet.Inputs
            self.name = self.worksheet.Name
            self.working_directory = self.worksheet.WorksheetWorkingDirectory
            #if filepath is None:  # no filepath value is passed when instancing

        def activate(self):
            """ activates the worksheet object """
            self.worksheet.activate

        def set_real_input(self, input_alias, value, units):
            """ Set the value of a numerical input range in the worksheet """
            if input_alias in self.worksheet.Inputs:
                self.worksheet.SetRealValue(str(input_alias),
                                              value,
                                              str(units))


    def activate(self):
        """ Activate the Mathcad window. If visible, this maximises Mathcad"""
        self.mcadapp.Activate()

    def active_sheet(self):
        """ Returns the active worksheet name """
        name = self.mcadapp.ActiveWorksheet.FullName
        if name == "":
            return None  # Returns none if the current worksheet not saved
        else:
            return name

    def worksheets(self):
        """ lists worksheets open in the Mathcad instance """
        self.mcadapp

    def close_all(self, save_option="Discard"):
        """ Closes all worksheets. Can specify save options before closing """
        if save_option in ["Discard", 2]:
            self.mcadapp.CloseAll(2)
        elif save_option in ["Prompt", 1]:
            self.mcadapp.CloseAll(1)
        elif save_option in ["Save", 0]:
            self.mcadapp.CloseAll(0)
        else:
            print("incorrect save argument specified")



if __name__ == "__main__":

    import os

    test = os.path.join(os.getcwd(),"Test","test.mcdx")

    print("Example usage")
    MC = Mathcad(visible=False)  # Open Mathcad with no GUI
    print(MC.active_sheet())  # Print the name of the active sheet
    MC.activate()
    print(MC.version)
    WS = MC.Worksheet(test)
    print(WS.name)
    print(WS.inputs)