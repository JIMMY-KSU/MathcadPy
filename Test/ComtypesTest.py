# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead
"""

import comtypes.client as CC
import Mathcad_Automation as MC


#tlbpath = r"C:\Program Files\PTC\Mathcad Prime 3.1\Ptc.MathcadPrime.Automation.tlb"
#mcad = CC.GetModule(tlbpath)
#print(mcad)
#
#MC.Application.

a = MC.Application()
a.Visible = True