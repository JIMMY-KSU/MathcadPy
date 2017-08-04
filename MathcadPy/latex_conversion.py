# -*- coding: utf-8 -*-
"""
MathcadPy
|
|- latex_conversion.py

Author: MattWoodhead

Requirements:

Mathcad Prime
comtypes (https://github.com/enthought/comtypes)

"""

import xml.etree.ElementTree as XMLET
from collections import namedtuple


# Unicode maths symbols and their latex equivalents
# TODO check the compatibility of these with latex mathmode!!
#__symbols = {'&amp;': r'\&',
#             'π': r'\pi ',
#             'α': r'\alpha ',
#             'β': r'\beta',
#             'γ': r'gamma',
#             '': r'\epsilon ',  # silly latex
#             'ε': r'\varepsilon ',  # This is epsilon!!
#             'φ': r'\phi ',
#             'θ': r'\theta ',
#             'ρ': r'\rho ',
#             'µ': r'\mu ',
#             '∆': r'\Delta ',
#             'ϕ': r'\Phi ',
#             '⇕': r'\Updownarrow ',
#             '⇔': r'\Leftrightarrow ',
#             'ω': r'\omega',
#             'Ω': r'\Omega',
#             '&': r'\&'
#             }
#
#symbols = namedtuple("Units", __symbols.keys())(**__symbols)

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

def read_mcdx(filepath):
    mcf = XMLET.parse(filepath).getroot()
    print(mcf)
    tags = {"def" : "http://schemas.mathsoft.com/worksheet30",
            "ml" : "http://schemas.mathsoft.com/math30",}
    regions = mcf.find("def:regions", tags)

    for region in regions:
        print(f"parsing region: {region.get('region-id')}")
        regiontype = region[0].tag[-4:]
        print(regiontype)

if __name__ == "__main__":
    file = r"C:\Users\Matt\Documents\GitHub\MathcadPy\Test\Layout_test\mathcad\worksheet.xml"
    read_mcdx(file)

