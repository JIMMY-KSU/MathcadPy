# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead

Requirements:

Mathcad Prime
Python Win 32 extensions (https://sourceforge.net/projects/pywin32/)

"""

import os

try:  # Check that dependencies are importable
    import comtypes.client as CC
except:
    print("The comtypes module is required")
    quit()  # Stop script


class Mathcad(object):
    """ Mathcad application object """
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
        for i in range(self.__mcadapp.Worksheets.Count):  # no. of open sheets
            worksheets.append(self.__mcadapp.Worksheets.Item(i).FullName)
        return worksheets  # Returns a list of open worksheet filenames

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

class Worksheet(object):
    """ Mathcad Worksheet object

    Either a filepath for a mathcad file can be supplied, or the
    filepath can be set to None (or similar) and the optional
    open_sheet_name argument can be used
    """

    def __init__(self, filepath, open_sheet_name=None):
        self.__mcadapp = CC.CreateObject("MathcadPrime.Application")
        self.__ws_at_init = {}
        for i in range(self.__mcadapp.Worksheets.Count):
            self.__ws_at_init[self.__mcadapp.Worksheets.Item(i).Name] = self.__mcadapp.Worksheets.Item(i).FullName
        if open_sheet_name is not None:
            for n, path in self.__ws_at_init.items():
                if open_sheet_name == n:
                    self.__mcadapp.Open(path)
                    self.__obj = self.__mcadapp.ActiveWorksheet  # Fetches COM worksheet object
                    self.Name = self.__obj.Name
                    break
            else:
                print(f"open_sheet_name={open_sheet_name} does not match the name of any open worksheets")
        if filepath is not None:
            if os.path.isfile(filepath) and os.path.exists(filepath):
                try:
                    self.__mcadapp.Open(filepath)
                    self.__obj = self.__mcadapp.ActiveWorksheet
                except:
                    print(f"Error opening {filepath}")

    def name(self):
        """ Returns the filename of the Worksheet object """
        return self.__obj.Name

    def inputs(self):
        """ returns a list of the designated input fields in the worksheet """
        _inputs = []
        for i in range(self.__obj.Inputs.Count):  # no. of open sheets
            _inputs.append(self.__obj.Inputs.GetAliasByIndex(i))
        return _inputs  # Returns a list of open worksheet filenames

    def outputs(self):
        """ returns a list of the designated output fields in the worksheet """
        outputs = []
        for i in range(self.__obj.Outputs.Count):
            outputs.append(self.__obj.Outputs.GetAliasByIndex(i))
        return outputs  # Returns a list of open worksheet filenames

    def activate(self):
        """ activates the worksheet object """
        self.__obj.Activate()

    def Close(self, save_option="Save"):
        """ Closes the worksheet """
        if save_option in ["Discard", 2]:
            self.__obj.Close(2)
        elif save_option in ["Prompt", 1]:
            self.__obj.Close(1)
        elif save_option in ["Save", 0]:
            self.__obj.Close(0)
        else:
            print("incorrect save argument specified")

    def save_as(self):
        pass
        self.Name = self.__mcadapp.ActiveWorksheet.Name

    def set_real_input(self, input_alias, value, units):
        """ Set the value of a numerical input range in the worksheet """
        if input_alias in self.__obj.Inputs:
            self.__obj.SetRealValue(str(input_alias), value, str(units))


if __name__ == "__main__":

    TEST = os.path.join(os.getcwd(), "Test", "test.mcdx")
    MC = Mathcad(visible=True) # Open Mathcad with no GUI
    MC.activate()
    WS = Worksheet(TEST)
    print(WS.inputs())
    print(WS.outputs())
