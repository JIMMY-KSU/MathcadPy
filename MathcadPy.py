# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead

Requirements:

Mathcad Prime
Python Win 32 extensions (https://sourceforge.net/projects/pywin32/)

"""

import win32com.client as win32

global PROGID
PROGID = "MathcadPrime.Application"

class Mathcad(object):
    """ top level wapper class """
    def __init__(self, visible=False):
        try:
            win32.gencache.EnsureModule("MathcadPrime.Application", 0, 1, 2)
            self.mcad = win32.Dispatch("MathcadPrime.Application")
            self.version = self.mcad.GetVersion()  # Fetches Mathcad version
            if visible in [True, False]:
                self.mcad.Visible=visible
            # @TODO mcad visible/hidden
        except:
            print("Could not locate the Mathcad Automation API")

    class Worksheet(object):
        """ Worksheet class """
        # @TODO
        pass

    def activate(self):
        """ Activate the Mathcad window. If visible, this maximises Mathcad"""
        self.mcad.Activate()

    def active_sheet(self):
        """ Returns the active worksheet name """
        return self.mcad.ActiveWorksheet

    def close_all(self, save_option="Discard"):
        if save_option in ["Discard", 2]:
            self.mcad.CloseAll(2)
        elif save_option in ["Prompt", 1]:
            self.mcad.CloseAll(1)
        elif save_option in ["Save", 0]:
            self.mcad.CloseAll(0)
        else:
            print("incorrect save argument specified")



if __name__ == "__main__":

    print("Example usage")
    MC = Mathcad()
    print(MC.active_sheet())
    MC.activate()
    print(MC.version)
