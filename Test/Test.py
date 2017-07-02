# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead
"""
import numpy as np

a = np.array([[1, 2], [3, 4], [5, 6]])

height, width = a.shape

print(isinstance(a, np.ndarray))

print(f"Height = {height}\nWidth = {width}")

