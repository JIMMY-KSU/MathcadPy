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
import xml.etree.ElementTree as ET


class _MathcadFile(object):
    """
    Class representing a .mcdx file.

    It can open and edit existing mathcad files.
    TODO - Create files from scratch
    """
    def __init__(self, filepath=None):
        self.filepath = filepath

        self.internal_files = None

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
            self.__zip_obj

    def _read_mcdx(self, filepath):
        """
        Uses zipfile to extract file information from the mcdx archive
        """
        try:
            iszip = zf.is_zipfile(filepath)
            extension = pathlib.PurePath(filepath).suffix
            if iszip and extension.lower() == ".mcdx":
                self.__zip_obj = zf.ZipFile(filepath)
                self.internal_files = {os.path.basename(i): i for i
                                       in self.__zip_obj.namelist()}
            else:
                raise TypeError("This module can only open .mcdx files")
                return False
        except IOError:
            print("Incorrect filepath")
            return False

    def _extract_file(self, filename):
        """
        Extracts the file with a matching filename from the archive
        incorrect or non existent filenames rais a KeyError
        """
        if self.internal_files is None:
            return False
        else:
            if filename in self.internal_files.keys():
                #return self.__zip_obj.read(filename)  # Returns bytes of file
                with self.__zip_obj.open(filename, 'r') as file:
                    return file  # Returns bytes of file
            else:
                raise KeyError("No matching file exists inside the archive")


if __name__ == "__main__":

    testpath = r"C:\Users\Matt\Documents\GitHub\MathcadPy\Test\Layout_test.mcdx"
    mcad_file = _MathcadFile(testpath)
    ws = mcad_file.internal_files["worksheet.xml"]
    ws_xml = mcad_file._extract_file(ws)
    #print(ET.parse(ws_xml))
