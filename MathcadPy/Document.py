# -*- coding: utf-8 -*-
"""
MathcadPy
|
|- Application.py

Author: MattWoodhead

Requirements:

Mathcad Prime
comtypes (https://github.com/enthought/comtypes)

"""

import zipfile as zf
import pathlib
import os
import xml.etree.ElementTree as XMLET
from collections import namedtuple

# Unicode maths symbols and their xml equivalents
# TODO check these are correct!
__symbols = {'&amp;': r'\&',
             'π': r'\pi ',
             'α': r'\alpha ',
             'β': r'\beta',
             'γ': r'gamma',
             '': r'\epsilon ',  # silly latex
             'ε': r'\varepsilon ',  # This is epsilon!!
             'φ': r'\phi ',
             'θ': r'\theta ',
             'ρ': r'\rho ',
             'µ': r'\mu ',
             '∆': r'\Delta ',
             'ϕ': r'\Phi ',
             '⇕': r'\Updownarrow ',
             '⇔': r'\Leftrightarrow ',
             'ω': r'\omega',
             'Ω': r'\Omega',
             '&': r'\&'
             }

symbols = namedtuple("Units", __symbols.keys())(**__symbols)

__functions = { "sin" : r"\sin",
               "cos" : r"\cos",
               "tan" : r"\tan",
               "cot" : r"\cot",
               }

functions = namedtuple("Units", __functions.keys())(**__functions)

__units = {"millimeter" : "mm",
    	   "meter"	: "m",
          "seconds" : "s",
    	   "minutes" : "min",
    	   "hours" : "h",
    	   "kilogram" : "kg",
    	   "newton" : "N"
          }

units = namedtuple("Units", __units.keys())(**__units)

class _MathcadFile(object):
    """
    Class representing a .mcdx file.

    It can open and edit existing mathcad files.
    TODO - Create files from scratch
    """
    def __init__(self, filepath=None):
        self.filepath = filepath

        self.internal_files = None
        self._zip_obj = None

        # Check to see if a filename has been passed when creating ther class
        if filepath == None:
            self.filepath == None
        elif filepath != None and not pathlib.Path(filepath).is_file():
            # Error handling for incorrect or non existent filename
            raise IOError("The file does not exist\n'{}'\n".format(filepath) +
                          "to create a new file use document()")
            self.filepath == None
        else:
            # Otherwise read in the file data to the class attributes
            self.name = pathlib.Path(filepath).stem  # No file extension
            self.filename = "{}.mcdx".format(self.name)
            self._read_mcdx(self.filepath)

    def _read_mcdx(self, filepath):
        """
        Uses zipfile to extract file information from the mcdx archive
        """
        try:
            iszip = zf.is_zipfile(filepath)
            extension = pathlib.PurePath(filepath).suffix
            if iszip and extension.lower() == ".mcdx":
                self._zip_obj = zf.ZipFile(filepath)
                z = self._zip_obj
                self.internal_files = {os.path.basename(i): z.read(i) for i
                                       in self._zip_obj.namelist()}

            else:
                raise TypeError("This module can only open .mcdx files")
                return False
        except IOError:
            print("Incorrect filepath")
            return False


if __name__ == "__main__":

    testpath = r"C:\Users\Matt\Documents\GitHub\MathcadPy\Test\Layout_test.mcdx"
    mcad_file = _MathcadFile(testpath)
    ws = mcad_file.internal_files["worksheet.xml"]
    print(ws)
    ws_xml = XMLET.fromstring(ws)
    print(ws_xml)

