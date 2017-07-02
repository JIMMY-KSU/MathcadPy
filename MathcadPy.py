# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead

Requirements:

Mathcad Prime
comtypes (https://github.com/enthought/comtypes)

"""

import os

try:  # Check that dependencies are importable
    import comtypes.client as CC
    import numpy as np
except:
    print("Not all required dependencies are installed")
    print("comtypes and numpy are required")
    quit()  # Stop script


class Mathcad(object):
    """ Mathcad application object """
    def __init__(self, visible=False):
        print("Loading Mathcad")
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

    # ~~~~~~~~~~~~~~~~~~~~~~~ File Operations ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
        """ Saves the worksheet under a new filename """
        pass
        self.Name = self.__mcadapp.ActiveWorksheet.Name

    def name(self):
        """ Returns the filename of the Worksheet object """
        return self.__obj.Name

    def readonly(self, setreadonly=None):
        """ Returns (and can optionally set) the worksheets read only status """
        if setreadonly is True:  # If readonly has been set to True
            self.__obj.IsReadOnly = True
        elif setreadonly is False: # If readonly has been set to False
            self.__obj.IsReadOnly = False
        return self.__obj.IsReadOnly  # Always return state

    def modified(self, setmodfied=None):
        """ Returns (and can optionally set) the worksheets modified status """
        if setmodfied is True:  # If readonly has been set to True
            self.__obj.Modified = True
        elif setmodfied is False: # If readonly has been set to False
            self.__obj.Modified = False
        return self.__obj.Modified  # Always return state

    # ~~~~~~~~~~~~~~~~~~~~~ Worksheet Operations ~~~~~~~~~~~~~~~~~~~~~~~~~~~

    def pause_calculation(self):
        """ Pauses worksheet calculation """
        self.__obj.PauseCalculation()

    def resume_calculation(self):
        """ Resumes the worksheets calculation """
        self.__obj.ResumeCalculation()

    def inputs(self):
        """ returns a list of the designated input fields in the worksheet """
        _inputs = []
        for i in range(self.__obj.Inputs.Count):  # no. of open sheets
            _inputs.append(self.__obj.Inputs.GetAliasByIndex(i))
        return _inputs  # Returns a list of open worksheet filenames

    def outputs(self):
        """ returns a list of the designated output fields in the worksheet """
        _outputs = []
        for i in range(self.__obj.Outputs.Count):
            _outputs.append(self.__obj.Outputs.GetAliasByIndex(i))
        return _outputs  # Returns a list of open worksheet filenames

    def create_matrix(self, rows, cols):
        """ Creates an empty Mathcad matrix of dimensions cols*rows """
        matrix_object = Matrix(rows, cols)

        return matrix_object

    def numpy_array_as_matrix(self, numpy_array):
        """ Takes a numpy array, creates a matrix, and populates the values """
        if isinstance(a, np.ndarray):
            height, width = numpy_array.shape  # Get array dimensions
            matrix = Matrix(height, width)

        else:
            print("Argument is not a Numpy array")

    def set_real_input(self, input_alias, value, units=""):
        """ Set the value of a numerical input range in the worksheet """
        if input_alias in self.inputs():  # Use inputs function to get list
            error = self.__obj.SetRealValue(str(input_alias), value, str(units))
            # COM command returns error count. 0 = everything set correctly
        else:
            print(f"{input_alias} is not a designated input field.\n\n" +
                  f"Available Input fields:\n{self.inputs()}")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' value/units\n" +
                  f"Check the '{self.__mcadapp.ActiveWorksheet.Name}' worksheet\n")
        return error

    def set_string_input(self, input_alias, string):
        """ Set the value of a numerical input range in the worksheet """
        if input_alias in self.inputs():  # Use inputs function to get list
            error = self.__obj.SetStringValue(str(input_alias), str(string))
            # COM command returns error count. 0 = everything set correctly
        else:
            print(f"{input_alias} is not a designated input field.\n\n" +
                  f"Available Input fields:\n{self.inputs()}")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' string\n" +
                  f"Check the '{self.__mcadapp.ActiveWorksheet.Name}' worksheet\n")
        return error


class Matrix(object):
    """ Mathcad Matrix object container """
    # Keeps methods inside Matrix for OOP
    def __init__(self, python_name=""):
        self.__mcadapp = CC.CreateObject("MathcadPrime.Application")
        self.__ws = self.__mcadapp.ActiveWorksheet
        self.python_name = python_name  # Just for organisation in scripts
        self.object = None

    def create_matrix(self, rows, columns):
        """ Creates a Mathcad matrix """
        try:
            rows, columns = int(rows), int(columns)
            self.shape = (self.rows, self.columns)
            self.object = self.__ws.CreateMatrix(rows, columns)
            return self.object
        except ValueError:
            raise ValueError("Matrix dimensions must be integers")
        except:
            raise Exception("COM Error creating Mathcad matrix")

    def set_element(self, row_index, column_index, value):
        if self.object is not None:
            try:
                row, col = int(row_index), int(column_index)
                self.object.SetMatrixElement(row, col, value)
            except ValueError:
                raise ValueError("Matrix dimensions must be integers")
            except:
                raise Exception("COM Error setting element value")
        else:
            raise TypeError("Matrix must first be created")










if __name__ == "__main__":

    TEST = os.path.join(os.getcwd(), "Test", "test.mcdx")
    MC = Mathcad(visible=True) # Open Mathcad with no GUI
    WS = Worksheet(TEST)
    a = WS.set_real_input("in_test", 2, "mm")
    print(a)
    print(WS.is_readonly())
