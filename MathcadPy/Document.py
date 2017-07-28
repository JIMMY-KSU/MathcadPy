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
from pathlib import Path


def __open_mcdx(filepath):


class document(object):
    """
    Class representing a .mcdx file.

    It can open and edit existing mathcad files.
    TODO - Create files from scratch
    """
    def __init__(self, filename=None):
        if filename == None:
            self.filename == None
        elif filename != None and not Path(filename).is_file():
            raise IOError("The filename does not exist\n'{}'\n".format(filename) +
                          "to create a new file use document()")
            self.filename == None
        else:
            self.filename == filename

    def __open(self, filename=self.filename):
        if filename == None:
            return False
        else:

document("C:\Z")