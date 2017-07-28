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

import zipfile
import pathlib


def _open_mcdx(filepath):
    try:
        iszip = zipfile.is_zipfile(filepath)
        extension = pathlib.PurePath(filepath).suffix
        if iszip and extension.lower() == ".mcdx":
            print(zipfile.ZipFile(filepath).printdir())
        else:
            raise TypeError("This module can only open .mcdx files")
            return False
    except IOError:
        print("Incorrect filepath")
        return False


class document(object):
    """
    Class representing a .mcdx file.

    It can open and edit existing mathcad files.
    TODO - Create files from scratch
    """
    def __init__(self, filename=None):
        self.filename = filename
        if filename == None:
            self.filename == None
        elif filename != None and not pathlib.Path(filename).is_file():
            raise IOError("The filename does not exist\n'{}'\n".format(filename) +
                          "to create a new file use document()")
            self.filename == None
        else:
            _open_mcdx(self.filename)


if __name__ == "__main__":

    testpath = r"C:\Users\Matt\Documents\GitHub\MathcadPy\Test\Layout_test.mcdx"
    document(testpath)