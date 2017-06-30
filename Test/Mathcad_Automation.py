# -*- coding: mbcs -*-
typelib_path = 'C:\\Program Files\\PTC\\Mathcad Prime 3.1\\Ptc.MathcadPrime.Automation.tlb'
_lcid = 0 # change this if required
from ctypes import *
import comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0
from comtypes import GUID
from ctypes import HRESULT
from comtypes import BSTR
from comtypes import helpstring
from comtypes import COMMETHOD
from comtypes import dispid
from ctypes.wintypes import VARIANT_BOOL
from comtypes import CoClass
import comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4
from comtypes.automation import VARIANT
from comtypes.automation import _midlSAFEARRAY


class IMathcadPrimeSetValueResults(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{AB5A25C9-8DFA-4F90-8C9F-FC297F3EDDB2}')
    _idlflags_ = ['dual', 'oleautomation']
IMathcadPrimeSetValueResults._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Count',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2)], HRESULT, 'GetResultByIndex',
              ( ['in'], c_int, 'indexArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(3)], HRESULT, 'GetResultByAlias',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeSetValueResults implementation
##class IMathcadPrimeSetValueResults_Impl(object):
##    @property
##    def Count(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def GetResultByIndex(self, indexArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def GetResultByAlias(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##

class IMathcadPrimeWorksheet2(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A04DC5F4-5E16-4D82-AF5A-C8A4CA2EAF8C}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeInputs(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A317A530-6337-4309-8F1C-C155488EFE96}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeOutputs(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A354B202-81A1-4EC9-A920-1876641ACF26}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeInputResult(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A31F2589-4FAD-4CA4-AB33-49672071CF29}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeOutputResult(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A3740C2B-8863-420B-8AFC-086488EE7DBF}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeOutputResultAs(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A3D8B0C9-9468-4650-804F-BF6D08297DE4}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeMatrix(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A17D73E0-B985-470F-ABA9-D2298799A950}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeInputMatrixResult(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A318D148-5D1D-4DFE-A0DE-A4AE28A4F3C2}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeOutputMatrixResult(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A3C56254-B233-48CD-86F7-5D8CE19A7097}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeOutputMatrixResultAs(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A3E6D703-E828-4EBF-BE27-8AE915CC97C7}')
    _idlflags_ = ['dual', 'oleautomation']

# values for enumeration 'SaveOption'
SaveOption_spSaveChanges = 0
SaveOption_spPromptToSaveChanges = 1
SaveOption_spDiscardChanges = 2
SaveOption = c_int # enum
IMathcadPrimeWorksheet2._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Name',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'FullName',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'IsReadOnly',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'Modified',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(4), 'propput'], HRESULT, 'Modified',
              ( ['in'], VARIANT_BOOL, 'pRetVal' )),
    COMMETHOD([dispid(5)], HRESULT, 'SetTitle',
              ( ['in'], BSTR, 'titleArg' )),
    COMMETHOD([dispid(6)], HRESULT, 'Save'),
    COMMETHOD([dispid(7)], HRESULT, 'SaveAs',
              ( ['in'], BSTR, 'newDocumentPathArg' )),
    COMMETHOD([dispid(8)], HRESULT, 'Synchronize'),
    COMMETHOD([dispid(9)], HRESULT, 'PauseCalculation'),
    COMMETHOD([dispid(10)], HRESULT, 'ResumeCalculation'),
    COMMETHOD([dispid(11)], HRESULT, 'SetRealValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], c_double, 'valueArg' ),
              ( ['in'], BSTR, 'unitsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(12), 'propget'], HRESULT, 'Inputs',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeInputs)), 'pRetVal' )),
    COMMETHOD([dispid(13), 'propget'], HRESULT, 'Outputs',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputs)), 'pRetVal' )),
    COMMETHOD([dispid(14)], HRESULT, 'InputGetRealValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeInputResult)), 'pRetVal' )),
    COMMETHOD([dispid(15)], HRESULT, 'OutputGetRealValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputResult)), 'pRetVal' )),
    COMMETHOD([dispid(16)], HRESULT, 'OutputGetRealValueAs',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'unitsArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputResultAs)), 'pRetVal' )),
    COMMETHOD([dispid(17)], HRESULT, 'CreateMatrix',
              ( ['in'], c_int, 'rowsArg' ),
              ( ['in'], c_int, 'columnsArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeMatrix)), 'pRetVal' )),
    COMMETHOD([dispid(18)], HRESULT, 'SetMatrixValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], POINTER(IMathcadPrimeMatrix), 'valueArg' ),
              ( ['in'], BSTR, 'unitsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(19)], HRESULT, 'InputGetMatrixValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeInputMatrixResult)), 'pRetVal' )),
    COMMETHOD([dispid(20)], HRESULT, 'OutputGetMatrixValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputMatrixResult)), 'pRetVal' )),
    COMMETHOD([dispid(21)], HRESULT, 'OutputGetMatrixValueAs',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'unitsArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputMatrixResultAs)), 'pRetVal' )),
    COMMETHOD([dispid(22)], HRESULT, 'IsOpen',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(23)], HRESULT, 'Activate'),
    COMMETHOD([dispid(24)], HRESULT, 'Close',
              ( ['in'], SaveOption, 'saveOptionArg' )),
]
################################################################
## code template for IMathcadPrimeWorksheet2 implementation
##class IMathcadPrimeWorksheet2_Impl(object):
##    @property
##    def Name(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def FullName(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def IsReadOnly(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    Modified = property(_get, _set, doc = _set.__doc__)
##
##    def SetTitle(self, titleArg):
##        '-no docstring-'
##        #return 
##
##    def Save(self):
##        '-no docstring-'
##        #return 
##
##    def SaveAs(self, newDocumentPathArg):
##        '-no docstring-'
##        #return 
##
##    def Synchronize(self):
##        '-no docstring-'
##        #return 
##
##    def PauseCalculation(self):
##        '-no docstring-'
##        #return 
##
##    def ResumeCalculation(self):
##        '-no docstring-'
##        #return 
##
##    def SetRealValue(self, aliasArg, valueArg, unitsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Inputs(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Outputs(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def InputGetRealValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetRealValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetRealValueAs(self, aliasArg, unitsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def CreateMatrix(self, rowsArg, columnsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def SetMatrixValue(self, aliasArg, valueArg, unitsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def InputGetMatrixValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetMatrixValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetMatrixValueAs(self, aliasArg, unitsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def IsOpen(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def Activate(self):
##        '-no docstring-'
##        #return 
##
##    def Close(self, saveOptionArg):
##        '-no docstring-'
##        #return 
##

class IMathcadPrimeInputsOutputsStates(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A72DE956-7106-4921-BFE0-90FCBE0DCA5B}')
    _idlflags_ = ['dual', 'oleautomation']
IMathcadPrimeInputsOutputsStates._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'ScalarInputsCount',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'ScalarOutputsCount',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(3)], HRESULT, 'GetInputScalarStateByIndex',
              ( ['in'], c_int, 'indexArg' ),
              ( ['out'], POINTER(BSTR), 'aliasArg' ),
              ( ['out'], POINTER(c_double), 'valueArg' ),
              ( ['out'], POINTER(BSTR), 'unitsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(4)], HRESULT, 'GetOutputScalarStateByIndex',
              ( ['in'], c_int, 'indexArg' ),
              ( ['out'], POINTER(BSTR), 'aliasArg' ),
              ( ['out'], POINTER(c_double), 'valueArg' ),
              ( ['out'], POINTER(BSTR), 'unitsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(5)], HRESULT, 'GetInputScalarStateByAlias',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out'], POINTER(c_double), 'valueArg' ),
              ( ['out'], POINTER(BSTR), 'unitsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(6)], HRESULT, 'GetOutputScalarStateByAlias',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out'], POINTER(c_double), 'valueArg' ),
              ( ['out'], POINTER(BSTR), 'unitsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeInputsOutputsStates implementation
##class IMathcadPrimeInputsOutputsStates_Impl(object):
##    @property
##    def ScalarInputsCount(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def ScalarOutputsCount(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def GetInputScalarStateByIndex(self, indexArg):
##        '-no docstring-'
##        #return aliasArg, valueArg, unitsArg, pRetVal
##
##    def GetOutputScalarStateByIndex(self, indexArg):
##        '-no docstring-'
##        #return aliasArg, valueArg, unitsArg, pRetVal
##
##    def GetInputScalarStateByAlias(self, aliasArg):
##        '-no docstring-'
##        #return valueArg, unitsArg, pRetVal
##
##    def GetOutputScalarStateByAlias(self, aliasArg):
##        '-no docstring-'
##        #return valueArg, unitsArg, pRetVal
##

class InputResult(CoClass):
    'Mathcad Prime InputResult Object'
    _reg_clsid_ = GUID('{A119F7C8-3DCF-4ED6-9BF2-B3BDE3552D25}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
InputResult._com_interfaces_ = [IMathcadPrimeInputResult, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

class OutputMatrixResultAs(CoClass):
    'Mathcad Prime OutputMatrixResultAs Object'
    _reg_clsid_ = GUID('{A3F99C0E-0529-4208-AA1D-46788D80ACCA}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
OutputMatrixResultAs._com_interfaces_ = [IMathcadPrimeOutputMatrixResultAs, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

class OutputResult(CoClass):
    'Mathcad Prime OutputResult Object'
    _reg_clsid_ = GUID('{A1ED586B-591A-4DDF-926F-E9778CD3D6CC}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
OutputResult._com_interfaces_ = [IMathcadPrimeOutputResult, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

class GetValueResult(CoClass):
    'Mathcad Prime GetValueResult Object'
    _reg_clsid_ = GUID('{AFFC9F46-5CB1-43F6-A477-F4B9A91BD24A}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
class IMathcadPrimeGetValueResult(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A6EB629A-AE39-4902-AFA8-76E1EFE632B6}')
    _idlflags_ = ['dual', 'oleautomation']
GetValueResult._com_interfaces_ = [IMathcadPrimeGetValueResult, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

class Outputs(CoClass):
    'Mathcad Prime Outputs Object'
    _reg_clsid_ = GUID('{A1B3240C-A124-4F3B-AF1A-7D5B5634B3C3}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
Outputs._com_interfaces_ = [IMathcadPrimeOutputs, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

IMathcadPrimeOutputs._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Count',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2)], HRESULT, 'GetAliasByIndex',
              ( ['in'], c_int, 'indexArg' ),
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeOutputs implementation
##class IMathcadPrimeOutputs_Impl(object):
##    @property
##    def Count(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def GetAliasByIndex(self, indexArg):
##        '-no docstring-'
##        #return pRetVal
##

class OutputResultAs(CoClass):
    'Mathcad Prime OutputResultAs Object'
    _reg_clsid_ = GUID('{A3F7211D-6D33-48EC-AFC7-A5471CB0F555}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
OutputResultAs._com_interfaces_ = [IMathcadPrimeOutputResultAs, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]


# values for enumeration 'WorksheetReadonlyOptionNames'
WorksheetReadonlyOptionNames_FileLocationHistoryDisabled = 0
WorksheetReadonlyOptionNames_OperationsWithEnabledStateGeneration = 1
WorksheetReadonlyOptionNames_RequestToUpdateInputsEnabled = 2
WorksheetReadonlyOptionNames_CaseInsensitiveAliasComparisonEnabled = 3
WorksheetReadonlyOptionNames = c_int # enum
class Application(CoClass):
    'Mathcad Prime Application Object'
    _reg_clsid_ = GUID('{A00E8B95-D415-433F-A04E-D298A54A7BB7}')
    _idlflags_ = []
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
class IMathcadPrimeApplication(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A027C8B4-F77A-4B1E-BE41-8B76FC865F25}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeApplication2(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A297208B-C701-4A2A-85B9-FCEC8115F0C6}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeApplication3(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A010504B-2FE6-402E-AD27-E24A8DE5C467}')
    _idlflags_ = ['dual', 'oleautomation']
Application._com_interfaces_ = [IMathcadPrimeApplication3, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object, IMathcadPrimeApplication, IMathcadPrimeApplication2]

class IMathcadPrimeEvents2(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{A0ECA09F-83C8-4536-B841-A33D981FBAFA}')
    _idlflags_ = ['oleautomation']

# values for enumeration 'WorksheetOperations'
WorksheetOperations_None = 0
WorksheetOperations_Save = 1
WorksheetOperations = c_int # enum
class IMathcadPrimeInputsOutputsConflicts(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{AB9C0902-3570-4697-89B4-F3887C6E978F}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeValuesSetter(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A9E81270-B96E-4AD0-9037-9D4BFD11CCC1}')
    _idlflags_ = ['dual', 'oleautomation']
IMathcadPrimeEvents2._methods_ = [
    COMMETHOD([], HRESULT, 'OnWorksheetSaved',
              ( ['in'], BSTR, 'documentFullNameArg' )),
    COMMETHOD([], HRESULT, 'OnWorksheetClosed',
              ( ['in'], BSTR, 'documentFullNameArg' ),
              ( ['in'], BSTR, 'documentNameArg' )),
    COMMETHOD([], HRESULT, 'OnWorksheetModified',
              ( ['in'], BSTR, 'documentFullNameArg' ),
              ( ['in'], BSTR, 'documentNameArg' ),
              ( ['in'], VARIANT_BOOL, 'isModifiedArg' )),
    COMMETHOD([], HRESULT, 'OnWorksheetRenamed',
              ( ['in'], BSTR, 'previousFullNameArg' ),
              ( ['in'], BSTR, 'currentFullNameArg' ),
              ( ['in'], BSTR, 'previousDocNameArg' ),
              ( ['in'], BSTR, 'currentDocNameArg' )),
    COMMETHOD([], HRESULT, 'OnWorksheetInputsOutputsSelected',
              ( ['in'], BSTR, 'documentFullNameArg' ),
              ( ['in'], BSTR, 'documentNameArg' ),
              ( ['in'], POINTER(IMathcadPrimeInputs), 'inputsArg' ),
              ( ['in'], POINTER(IMathcadPrimeOutputs), 'outputsArg' )),
    COMMETHOD([], HRESULT, 'OnExit'),
    COMMETHOD([], HRESULT, 'OnWorksheetStatesGenerated',
              ( ['in'], BSTR, 'documentFullNameArg' ),
              ( ['in'], BSTR, 'documentNameArg' ),
              ( ['in'], WorksheetOperations, 'operationsArg' ),
              ( ['in'], POINTER(IMathcadPrimeInputsOutputsStates), 'itemsStatesArg' ),
              ( ['in'], POINTER(IMathcadPrimeInputsOutputsConflicts), 'conflictsArg' )),
    COMMETHOD([], HRESULT, 'OnWorksheetStatesGenerating',
              ( ['in'], BSTR, 'documentFullNameArg' ),
              ( ['in'], BSTR, 'documentNameArg' ),
              ( ['in'], WorksheetOperations, 'operationsArg' ),
              ( ['in'], POINTER(IMathcadPrimeInputsOutputsStates), 'itemsStatesArg' ),
              ( ['in'], POINTER(IMathcadPrimeInputsOutputsConflicts), 'conflictsArg' )),
    COMMETHOD([], HRESULT, 'OnWorksheetRequestToUpdateInputs',
              ( ['in'], BSTR, 'documentFullNameArg' ),
              ( ['in'], BSTR, 'documentNameArg' ),
              ( ['in'], POINTER(IMathcadPrimeValuesSetter), 'setterArg' )),
]
################################################################
## code template for IMathcadPrimeEvents2 implementation
##class IMathcadPrimeEvents2_Impl(object):
##    def OnWorksheetSaved(self, documentFullNameArg):
##        '-no docstring-'
##        #return 
##
##    def OnWorksheetClosed(self, documentFullNameArg, documentNameArg):
##        '-no docstring-'
##        #return 
##
##    def OnWorksheetModified(self, documentFullNameArg, documentNameArg, isModifiedArg):
##        '-no docstring-'
##        #return 
##
##    def OnWorksheetRenamed(self, previousFullNameArg, currentFullNameArg, previousDocNameArg, currentDocNameArg):
##        '-no docstring-'
##        #return 
##
##    def OnWorksheetInputsOutputsSelected(self, documentFullNameArg, documentNameArg, inputsArg, outputsArg):
##        '-no docstring-'
##        #return 
##
##    def OnExit(self):
##        '-no docstring-'
##        #return 
##
##    def OnWorksheetStatesGenerated(self, documentFullNameArg, documentNameArg, operationsArg, itemsStatesArg, conflictsArg):
##        '-no docstring-'
##        #return 
##
##    def OnWorksheetStatesGenerating(self, documentFullNameArg, documentNameArg, operationsArg, itemsStatesArg, conflictsArg):
##        '-no docstring-'
##        #return 
##
##    def OnWorksheetRequestToUpdateInputs(self, documentFullNameArg, documentNameArg, setterArg):
##        '-no docstring-'
##        #return 
##

class ApplicationObsolete(CoClass):
    'Mathcad Prime ApplicationObsolete Object'
    _reg_clsid_ = GUID('{A3E4E622-5CFF-4973-96EB-1E3EAB5151C8}')
    _idlflags_ = []
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
ApplicationObsolete._com_interfaces_ = [IMathcadPrimeApplication2, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object, IMathcadPrimeApplication]

class ValuesSetter(CoClass):
    'Mathcad Prime ValuesSetter Object'
    _reg_clsid_ = GUID('{A666719D-79F4-449B-A4C1-8DEECD5FA18A}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
ValuesSetter._com_interfaces_ = [IMathcadPrimeValuesSetter, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

class Inputs(CoClass):
    'Mathcad Prime Inputs Object'
    _reg_clsid_ = GUID('{A2E0BC48-0B59-4703-923D-97B59069C511}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
Inputs._com_interfaces_ = [IMathcadPrimeInputs, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

IMathcadPrimeOutputResult._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'ErrorCode',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'RealResult',
              ( ['out', 'retval'], POINTER(c_double), 'pRetVal' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'Units',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeOutputResult implementation
##class IMathcadPrimeOutputResult_Impl(object):
##    @property
##    def ErrorCode(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def RealResult(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Units(self):
##        '-no docstring-'
##        #return pRetVal
##

IMathcadPrimeInputsOutputsConflicts._methods_ = [
    COMMETHOD([dispid(1)], HRESULT, 'AddGeneralWarning',
              ( ['in'], BSTR, 'warningArg' )),
    COMMETHOD([dispid(2)], HRESULT, 'AddGeneralError',
              ( ['in'], BSTR, 'errorArg' )),
    COMMETHOD([dispid(3)], HRESULT, 'AddItemWarning',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'warningArg' )),
    COMMETHOD([dispid(4)], HRESULT, 'AddItemError',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'errorArg' )),
]
################################################################
## code template for IMathcadPrimeInputsOutputsConflicts implementation
##class IMathcadPrimeInputsOutputsConflicts_Impl(object):
##    def AddGeneralWarning(self, warningArg):
##        '-no docstring-'
##        #return 
##
##    def AddGeneralError(self, errorArg):
##        '-no docstring-'
##        #return 
##
##    def AddItemWarning(self, aliasArg, warningArg):
##        '-no docstring-'
##        #return 
##
##    def AddItemError(self, aliasArg, errorArg):
##        '-no docstring-'
##        #return 
##


# values for enumeration 'MathcadPrimeEvents'
MathcadPrimeEvents_OnExit = 100
MathcadPrimeEvents_OnWorksheetSaved = 0
MathcadPrimeEvents_OnWorksheetClosed = 1
MathcadPrimeEvents_OnWorksheetModified = 2
MathcadPrimeEvents_OnWorksheetRenamed = 3
MathcadPrimeEvents_OnWorksheetInputsOutputsSelected = 4
MathcadPrimeEvents_OnWorksheetStatesGenerated = 5
MathcadPrimeEvents_OnWorksheetStatesGenerating = 6
MathcadPrimeEvents_OnRequestToUpdateInputs = 7
MathcadPrimeEvents = c_int # enum
class InputsOutputsConflicts(CoClass):
    'Mathcad Prime InputsOutputsConflicts Object'
    _reg_clsid_ = GUID('{A2FEB23B-CDAE-4A1A-BFFE-30194E4C618D}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
InputsOutputsConflicts._com_interfaces_ = [IMathcadPrimeInputsOutputsConflicts, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

class Worksheet(CoClass):
    'Mathcad Prime Worksheet Object'
    _reg_clsid_ = GUID('{A2C7A48C-4B32-495E-AF0B-8357B115A48C}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
class IMathcadPrimeWorksheet(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A17F8C1D-A683-488D-AE43-3B0860FB5B2F}')
    _idlflags_ = ['dual', 'oleautomation']
class IMathcadPrimeWorksheet3(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A27AD87C-6F4E-4B8A-8827-06A22ED16F35}')
    _idlflags_ = ['dual', 'oleautomation']
Worksheet._com_interfaces_ = [IMathcadPrimeWorksheet3, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object, IMathcadPrimeWorksheet, IMathcadPrimeWorksheet2]

class IMathcadPrimeEvents(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{A170C6A4-3DEB-43A7-A5C4-9164EF85D1C2}')
    _idlflags_ = ['oleautomation']
IMathcadPrimeEvents._methods_ = [
    COMMETHOD([], HRESULT, 'OnSelect',
              ( ['in'], POINTER(IMathcadPrimeInputs), 'inputsOnSelectArg' ),
              ( ['in'], POINTER(IMathcadPrimeOutputs), 'outputsOnSelectArg' )),
    COMMETHOD([], HRESULT, 'OnSave',
              ( ['in'], BSTR, 'documentNameArg' )),
    COMMETHOD([], HRESULT, 'OnExit'),
]
################################################################
## code template for IMathcadPrimeEvents implementation
##class IMathcadPrimeEvents_Impl(object):
##    def OnSelect(self, inputsOnSelectArg, outputsOnSelectArg):
##        '-no docstring-'
##        #return 
##
##    def OnSave(self, documentNameArg):
##        '-no docstring-'
##        #return 
##
##    def OnExit(self):
##        '-no docstring-'
##        #return 
##


# values for enumeration 'ValueResultTypes'
ValueResultTypes_None = 0
ValueResultTypes_Real = 1
ValueResultTypes_String = 2
ValueResultTypes_Matrix = 3
ValueResultTypes = c_int # enum
IMathcadPrimeGetValueResult._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'ResultType',
              ( ['out', 'retval'], POINTER(ValueResultTypes), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'ErrorCode',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'Units',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'RealResult',
              ( ['out', 'retval'], POINTER(c_double), 'pRetVal' )),
    COMMETHOD([dispid(5), 'propget'], HRESULT, 'StringResult',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(6), 'propget'], HRESULT, 'MatrixResult',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeMatrix)), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeGetValueResult implementation
##class IMathcadPrimeGetValueResult_Impl(object):
##    @property
##    def ResultType(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def ErrorCode(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Units(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def RealResult(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def StringResult(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def MatrixResult(self):
##        '-no docstring-'
##        #return pRetVal
##

IMathcadPrimeMatrix._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Rows',
              ( ['out', 'retval'], POINTER(c_double), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'Columns',
              ( ['out', 'retval'], POINTER(c_double), 'pRetVal' )),
    COMMETHOD([dispid(3)], HRESULT, 'SetMatrixElement',
              ( ['in'], c_int, 'rowIndexArg' ),
              ( ['in'], c_int, 'columnIndexArg' ),
              ( ['in'], c_double, 'valueArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(4)], HRESULT, 'GetMatrixElement',
              ( ['in'], c_int, 'rowIndexArg' ),
              ( ['in'], c_int, 'columnIndexArg' ),
              ( ['out'], POINTER(c_double), 'valueArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeMatrix implementation
##class IMathcadPrimeMatrix_Impl(object):
##    @property
##    def Rows(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Columns(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def SetMatrixElement(self, rowIndexArg, columnIndexArg, valueArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def GetMatrixElement(self, rowIndexArg, columnIndexArg):
##        '-no docstring-'
##        #return valueArg, pRetVal
##

IMathcadPrimeApplication._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Visible',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(1), 'propput'], HRESULT, 'Visible',
              ( ['in'], VARIANT_BOOL, 'pRetVal' )),
    COMMETHOD([dispid(2)], HRESULT, 'Activate'),
    COMMETHOD([dispid(3)], HRESULT, 'Quit',
              ( ['in'], SaveOption, 'saveOptionArg' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'ActiveWorksheet',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheet)), 'pRetVal' )),
    COMMETHOD([dispid(5)], HRESULT, 'Open',
              ( ['in'], BSTR, 'documentPathArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheet)), 'pRetVal' )),
    COMMETHOD([dispid(6)], HRESULT, 'InitializeEvents',
              ( ['in'], POINTER(IMathcadPrimeEvents), 'eventsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeApplication implementation
##class IMathcadPrimeApplication_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    Visible = property(_get, _set, doc = _set.__doc__)
##
##    def Activate(self):
##        '-no docstring-'
##        #return 
##
##    def Quit(self, saveOptionArg):
##        '-no docstring-'
##        #return 
##
##    @property
##    def ActiveWorksheet(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def Open(self, documentPathArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def InitializeEvents(self, eventsArg):
##        '-no docstring-'
##        #return pRetVal
##

IMathcadPrimeOutputMatrixResultAs._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'ErrorCode',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'MatrixResult',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeMatrix)), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeOutputMatrixResultAs implementation
##class IMathcadPrimeOutputMatrixResultAs_Impl(object):
##    @property
##    def ErrorCode(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def MatrixResult(self):
##        '-no docstring-'
##        #return pRetVal
##

IMathcadPrimeValuesSetter._methods_ = [
    COMMETHOD([dispid(1)], HRESULT, 'AddScalarValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], c_double, 'valueArg' ),
              ( ['in'], BSTR, 'unitsArg' )),
    COMMETHOD([dispid(2)], HRESULT, 'AddMatrixValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], _midlSAFEARRAY(VARIANT), 'valueArg' ),
              ( ['in'], BSTR, 'unitsArg' )),
    COMMETHOD([dispid(3)], HRESULT, 'AddStringValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'valueArg' )),
    COMMETHOD([dispid(4)], HRESULT, 'AddSExprValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'sexpressionArg' )),
    COMMETHOD([dispid(5)], HRESULT, 'SetValues',
              ( ['in'], c_int, 'secondsArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeSetValueResults)), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeValuesSetter implementation
##class IMathcadPrimeValuesSetter_Impl(object):
##    def AddScalarValue(self, aliasArg, valueArg, unitsArg):
##        '-no docstring-'
##        #return 
##
##    def AddMatrixValue(self, aliasArg, valueArg, unitsArg):
##        '-no docstring-'
##        #return 
##
##    def AddStringValue(self, aliasArg, valueArg):
##        '-no docstring-'
##        #return 
##
##    def AddSExprValue(self, aliasArg, sexpressionArg):
##        '-no docstring-'
##        #return 
##
##    def SetValues(self, secondsArg):
##        '-no docstring-'
##        #return pRetVal
##

IMathcadPrimeInputMatrixResult._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'ErrorCode',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'MatrixResult',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeMatrix)), 'pRetVal' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'Units',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeInputMatrixResult implementation
##class IMathcadPrimeInputMatrixResult_Impl(object):
##    @property
##    def ErrorCode(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def MatrixResult(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Units(self):
##        '-no docstring-'
##        #return pRetVal
##

class InputMatrixResult(CoClass):
    'Mathcad Prime InputMatrixResult Object'
    _reg_clsid_ = GUID('{A1FD7C0C-287C-4AA9-AF64-274575A06598}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
InputMatrixResult._com_interfaces_ = [IMathcadPrimeInputMatrixResult, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

class Worksheets(CoClass):
    'Mathcad Prime Worksheets Object'
    _reg_clsid_ = GUID('{A31DBEE4-5053-4A2C-832F-B0161995FAE8}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
class IMathcadPrimeWorksheets(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A2E041D6-4946-40AD-9AC9-56F0A2FFA3FF}')
    _idlflags_ = ['dual', 'oleautomation']
Worksheets._com_interfaces_ = [IMathcadPrimeWorksheets, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

class IMathcadPrimeWorksheetReadonlyOptions(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{A11258FF-2491-480C-9E3B-EBBF08AF1B72}')
    _idlflags_ = ['dual', 'oleautomation']
IMathcadPrimeWorksheetReadonlyOptions._methods_ = [
    COMMETHOD([dispid(1)], HRESULT, 'SetOptionValue',
              ( ['in'], WorksheetReadonlyOptionNames, 'optionNameArg' ),
              ( ['in'], VARIANT, 'optionValueArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeWorksheetReadonlyOptions implementation
##class IMathcadPrimeWorksheetReadonlyOptions_Impl(object):
##    def SetOptionValue(self, optionNameArg, optionValueArg):
##        '-no docstring-'
##        #return pRetVal
##

class WorksheetReadonlyOptions(CoClass):
    'Mathcad Prime WorksheetReadonlyOptions Object'
    _reg_clsid_ = GUID('{A0568616-45DB-4B89-87D6-6769D4AEFFE4}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
WorksheetReadonlyOptions._com_interfaces_ = [IMathcadPrimeWorksheetReadonlyOptions, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

IMathcadPrimeApplication2._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Visible',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(1), 'propput'], HRESULT, 'Visible',
              ( ['in'], VARIANT_BOOL, 'pRetVal' )),
    COMMETHOD([dispid(2)], HRESULT, 'Activate'),
    COMMETHOD([dispid(3)], HRESULT, 'Quit',
              ( ['in'], SaveOption, 'saveOptionArg' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'ActiveWorksheet',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheet)), 'pRetVal' )),
    COMMETHOD([dispid(5)], HRESULT, 'Open',
              ( ['in'], BSTR, 'documentPathArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheet)), 'pRetVal' )),
    COMMETHOD([dispid(6)], HRESULT, 'InitializeEvents',
              ( ['in'], POINTER(IMathcadPrimeEvents), 'eventsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(7), 'propget'], HRESULT, 'Worksheets',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheets)), 'pRetVal' )),
    COMMETHOD([dispid(8)], HRESULT, 'CloseAll',
              ( ['in'], SaveOption, 'saveOptionArg' )),
    COMMETHOD([dispid(9)], HRESULT, 'GetVersion',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeApplication2 implementation
##class IMathcadPrimeApplication2_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    Visible = property(_get, _set, doc = _set.__doc__)
##
##    def Activate(self):
##        '-no docstring-'
##        #return 
##
##    def Quit(self, saveOptionArg):
##        '-no docstring-'
##        #return 
##
##    @property
##    def ActiveWorksheet(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def Open(self, documentPathArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def InitializeEvents(self, eventsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Worksheets(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def CloseAll(self, saveOptionArg):
##        '-no docstring-'
##        #return 
##
##    def GetVersion(self):
##        '-no docstring-'
##        #return pRetVal
##

IMathcadPrimeWorksheet3._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Name',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'FullName',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'IsReadOnly',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'Modified',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(4), 'propput'], HRESULT, 'Modified',
              ( ['in'], VARIANT_BOOL, 'pRetVal' )),
    COMMETHOD([dispid(5)], HRESULT, 'SetTitle',
              ( ['in'], BSTR, 'titleArg' )),
    COMMETHOD([dispid(6)], HRESULT, 'Save'),
    COMMETHOD([dispid(7)], HRESULT, 'SaveAs',
              ( ['in'], BSTR, 'newDocumentPathArg' )),
    COMMETHOD([dispid(8)], HRESULT, 'Synchronize'),
    COMMETHOD([dispid(9)], HRESULT, 'PauseCalculation'),
    COMMETHOD([dispid(10)], HRESULT, 'ResumeCalculation'),
    COMMETHOD([dispid(11)], HRESULT, 'SetRealValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], c_double, 'valueArg' ),
              ( ['in'], BSTR, 'unitsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(12), 'propget'], HRESULT, 'Inputs',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeInputs)), 'pRetVal' )),
    COMMETHOD([dispid(13), 'propget'], HRESULT, 'Outputs',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputs)), 'pRetVal' )),
    COMMETHOD([dispid(14)], HRESULT, 'InputGetRealValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeInputResult)), 'pRetVal' )),
    COMMETHOD([dispid(15)], HRESULT, 'OutputGetRealValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputResult)), 'pRetVal' )),
    COMMETHOD([dispid(16)], HRESULT, 'OutputGetRealValueAs',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'unitsArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputResultAs)), 'pRetVal' )),
    COMMETHOD([dispid(17)], HRESULT, 'CreateMatrix',
              ( ['in'], c_int, 'rowsArg' ),
              ( ['in'], c_int, 'columnsArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeMatrix)), 'pRetVal' )),
    COMMETHOD([dispid(18)], HRESULT, 'SetMatrixValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], POINTER(IMathcadPrimeMatrix), 'valueArg' ),
              ( ['in'], BSTR, 'unitsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(19)], HRESULT, 'InputGetMatrixValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeInputMatrixResult)), 'pRetVal' )),
    COMMETHOD([dispid(20)], HRESULT, 'OutputGetMatrixValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputMatrixResult)), 'pRetVal' )),
    COMMETHOD([dispid(21)], HRESULT, 'OutputGetMatrixValueAs',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'unitsArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputMatrixResultAs)), 'pRetVal' )),
    COMMETHOD([dispid(22)], HRESULT, 'IsOpen',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(23)], HRESULT, 'Activate'),
    COMMETHOD([dispid(24)], HRESULT, 'Close',
              ( ['in'], SaveOption, 'saveOptionArg' )),
    COMMETHOD([dispid(25)], HRESULT, 'GetTitle',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(26), 'propget'], HRESULT, 'WorksheetTabIcon',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(26), 'propput'], HRESULT, 'WorksheetTabIcon',
              ( ['in'], BSTR, 'pRetVal' )),
    COMMETHOD([dispid(27), 'propget'], HRESULT, 'WorksheetTabName',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(27), 'propput'], HRESULT, 'WorksheetTabName',
              ( ['in'], BSTR, 'pRetVal' )),
    COMMETHOD([dispid(28), 'propget'], HRESULT, 'WorksheetClosingPrompt',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(28), 'propput'], HRESULT, 'WorksheetClosingPrompt',
              ( ['in'], BSTR, 'pRetVal' )),
    COMMETHOD([dispid(29), 'propget'], HRESULT, 'WorksheetDisplayedFilePath',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(29), 'propput'], HRESULT, 'WorksheetDisplayedFilePath',
              ( ['in'], BSTR, 'pRetVal' )),
    COMMETHOD([dispid(30), 'propget'], HRESULT, 'WorksheetWorkingDirectory',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(30), 'propput'], HRESULT, 'WorksheetWorkingDirectory',
              ( ['in'], BSTR, 'pRetVal' )),
    COMMETHOD([dispid(31)], HRESULT, 'GetWorksheetReadOnlyOptionValue',
              ( ['in'], WorksheetReadonlyOptionNames, 'optionNameArg' ),
              ( ['out', 'retval'], POINTER(VARIANT), 'pRetVal' )),
    COMMETHOD([dispid(32)], HRESULT, 'CreateValuesSetter',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeValuesSetter)), 'pRetVal' )),
    COMMETHOD([dispid(33)], HRESULT, 'SetStringValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'valueArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(34)], HRESULT, 'InputGetStringValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(35)], HRESULT, 'OutputGetStringValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(36)], HRESULT, 'SetSExprValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'sexpressionArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(37)], HRESULT, 'InputGetSExprValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(38)], HRESULT, 'InputGetValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeGetValueResult)), 'pRetVal' )),
    COMMETHOD([dispid(39)], HRESULT, 'OutputGetValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeGetValueResult)), 'pRetVal' )),
    COMMETHOD([dispid(40), 'propget'], HRESULT, 'DefaultCalculationTimeout',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(40), 'propput'], HRESULT, 'DefaultCalculationTimeout',
              ( ['in'], c_int, 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeWorksheet3 implementation
##class IMathcadPrimeWorksheet3_Impl(object):
##    @property
##    def Name(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def FullName(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def IsReadOnly(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    Modified = property(_get, _set, doc = _set.__doc__)
##
##    def SetTitle(self, titleArg):
##        '-no docstring-'
##        #return 
##
##    def Save(self):
##        '-no docstring-'
##        #return 
##
##    def SaveAs(self, newDocumentPathArg):
##        '-no docstring-'
##        #return 
##
##    def Synchronize(self):
##        '-no docstring-'
##        #return 
##
##    def PauseCalculation(self):
##        '-no docstring-'
##        #return 
##
##    def ResumeCalculation(self):
##        '-no docstring-'
##        #return 
##
##    def SetRealValue(self, aliasArg, valueArg, unitsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Inputs(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Outputs(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def InputGetRealValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetRealValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetRealValueAs(self, aliasArg, unitsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def CreateMatrix(self, rowsArg, columnsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def SetMatrixValue(self, aliasArg, valueArg, unitsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def InputGetMatrixValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetMatrixValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetMatrixValueAs(self, aliasArg, unitsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def IsOpen(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def Activate(self):
##        '-no docstring-'
##        #return 
##
##    def Close(self, saveOptionArg):
##        '-no docstring-'
##        #return 
##
##    def GetTitle(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    WorksheetTabIcon = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    WorksheetTabName = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    WorksheetClosingPrompt = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    WorksheetDisplayedFilePath = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    WorksheetWorkingDirectory = property(_get, _set, doc = _set.__doc__)
##
##    def GetWorksheetReadOnlyOptionValue(self, optionNameArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def CreateValuesSetter(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def SetStringValue(self, aliasArg, valueArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def InputGetStringValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetStringValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def SetSExprValue(self, aliasArg, sexpressionArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def InputGetSExprValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def InputGetValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    DefaultCalculationTimeout = property(_get, _set, doc = _set.__doc__)
##

class InputsOutputsStates(CoClass):
    'Mathcad Prime InputsOutputsStates Object'
    _reg_clsid_ = GUID('{A31B7E02-2FD3-4FEB-907C-505F321E2E03}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
InputsOutputsStates._com_interfaces_ = [IMathcadPrimeInputsOutputsStates, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

class OutputMatrixResult(CoClass):
    'Mathcad Prime OutputMatrixResult Object'
    _reg_clsid_ = GUID('{A16907E6-0B05-4D9C-A2E3-F47D9D3425A1}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
OutputMatrixResult._com_interfaces_ = [IMathcadPrimeOutputMatrixResult, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

class Library(object):
    'PTC Mathcad Prime COM Automation'
    name = 'Ptc_MathcadPrime_Automation'
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)

IMathcadPrimeWorksheets._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Count',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2)], HRESULT, 'Item',
              ( ['in'], c_int, 'indexArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheet2)), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeWorksheets implementation
##class IMathcadPrimeWorksheets_Impl(object):
##    @property
##    def Count(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def Item(self, indexArg):
##        '-no docstring-'
##        #return pRetVal
##

IMathcadPrimeApplication3._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Visible',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(1), 'propput'], HRESULT, 'Visible',
              ( ['in'], VARIANT_BOOL, 'pRetVal' )),
    COMMETHOD([dispid(2)], HRESULT, 'Activate'),
    COMMETHOD([dispid(3)], HRESULT, 'Quit',
              ( ['in'], SaveOption, 'saveOptionArg' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'ActiveWorksheet',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheet)), 'pRetVal' )),
    COMMETHOD([dispid(5)], HRESULT, 'Open',
              ( ['in'], BSTR, 'documentPathArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheet)), 'pRetVal' )),
    COMMETHOD([dispid(6)], HRESULT, 'InitializeEvents',
              ( ['in'], POINTER(IMathcadPrimeEvents), 'eventsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(7), 'propget'], HRESULT, 'Worksheets',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheets)), 'pRetVal' )),
    COMMETHOD([dispid(8)], HRESULT, 'CloseAll',
              ( ['in'], SaveOption, 'saveOptionArg' )),
    COMMETHOD([dispid(9)], HRESULT, 'GetVersion',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(10)], HRESULT, 'InitializeEvents2',
              ( ['in'], POINTER(IMathcadPrimeEvents2), 'eventsArg' ),
              ( ['in'], VARIANT_BOOL, 'subscribeAllArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(11)], HRESULT, 'SubscribeEvent',
              ( ['in'], MathcadPrimeEvents, 'primeEventArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(12)], HRESULT, 'UnsubscribeEvent',
              ( ['in'], MathcadPrimeEvents, 'primeEventArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(13)], HRESULT, 'CreateWorksheetReadonlyOptions',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheetReadonlyOptions)), 'pRetVal' )),
    COMMETHOD([dispid(14)], HRESULT, 'OpenEx',
              ( ['in'], BSTR, 'documentPathArg' ),
              ( ['in'], POINTER(IMathcadPrimeWorksheetReadonlyOptions), 'optionsArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeWorksheet3)), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeApplication3 implementation
##class IMathcadPrimeApplication3_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    Visible = property(_get, _set, doc = _set.__doc__)
##
##    def Activate(self):
##        '-no docstring-'
##        #return 
##
##    def Quit(self, saveOptionArg):
##        '-no docstring-'
##        #return 
##
##    @property
##    def ActiveWorksheet(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def Open(self, documentPathArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def InitializeEvents(self, eventsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Worksheets(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def CloseAll(self, saveOptionArg):
##        '-no docstring-'
##        #return 
##
##    def GetVersion(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def InitializeEvents2(self, eventsArg, subscribeAllArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def SubscribeEvent(self, primeEventArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def UnsubscribeEvent(self, primeEventArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def CreateWorksheetReadonlyOptions(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def OpenEx(self, documentPathArg, optionsArg):
##        '-no docstring-'
##        #return pRetVal
##

IMathcadPrimeInputResult._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'ErrorCode',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'RealResult',
              ( ['out', 'retval'], POINTER(c_double), 'pRetVal' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'Units',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeInputResult implementation
##class IMathcadPrimeInputResult_Impl(object):
##    @property
##    def ErrorCode(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def RealResult(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Units(self):
##        '-no docstring-'
##        #return pRetVal
##

class Matrix(CoClass):
    'Mathcad Prime Matrix Object'
    _reg_clsid_ = GUID('{A258774A-1F72-431C-9464-89F831C4CCAF}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
Matrix._com_interfaces_ = [IMathcadPrimeMatrix, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

IMathcadPrimeOutputResultAs._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'ErrorCode',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'RealResult',
              ( ['out', 'retval'], POINTER(c_double), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeOutputResultAs implementation
##class IMathcadPrimeOutputResultAs_Impl(object):
##    @property
##    def ErrorCode(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def RealResult(self):
##        '-no docstring-'
##        #return pRetVal
##

class SetValueResults(CoClass):
    'Mathcad Prime SetValueResults Object'
    _reg_clsid_ = GUID('{AFCA092C-04CB-44B4-8FCF-90E6F746B056}')
    _idlflags_ = ['noncreatable']
    _typelib_path_ = typelib_path
    _reg_typelib_ = ('{A24EB614-A183-400F-8207-1E58D61945D6}', 1, 0)
SetValueResults._com_interfaces_ = [IMathcadPrimeSetValueResults, comtypes.gen._BED7F4EA_1A96_11D2_8F08_00A0C9A6186D_0_2_4._Object]

IMathcadPrimeWorksheet._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Name',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'FullName',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'IsReadOnly',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'Modified',
              ( ['out', 'retval'], POINTER(VARIANT_BOOL), 'pRetVal' )),
    COMMETHOD([dispid(4), 'propput'], HRESULT, 'Modified',
              ( ['in'], VARIANT_BOOL, 'pRetVal' )),
    COMMETHOD([dispid(5)], HRESULT, 'SetTitle',
              ( ['in'], BSTR, 'titleArg' )),
    COMMETHOD([dispid(6)], HRESULT, 'Save'),
    COMMETHOD([dispid(7)], HRESULT, 'SaveAs',
              ( ['in'], BSTR, 'newDocumentPathArg' )),
    COMMETHOD([dispid(8)], HRESULT, 'Synchronize'),
    COMMETHOD([dispid(9)], HRESULT, 'PauseCalculation'),
    COMMETHOD([dispid(10)], HRESULT, 'ResumeCalculation'),
    COMMETHOD([dispid(11)], HRESULT, 'SetRealValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], c_double, 'valueArg' ),
              ( ['in'], BSTR, 'unitsArg' ),
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(12), 'propget'], HRESULT, 'Inputs',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeInputs)), 'pRetVal' )),
    COMMETHOD([dispid(13), 'propget'], HRESULT, 'Outputs',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputs)), 'pRetVal' )),
    COMMETHOD([dispid(14)], HRESULT, 'InputGetRealValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeInputResult)), 'pRetVal' )),
    COMMETHOD([dispid(15)], HRESULT, 'OutputGetRealValue',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputResult)), 'pRetVal' )),
    COMMETHOD([dispid(16)], HRESULT, 'OutputGetRealValueAs',
              ( ['in'], BSTR, 'aliasArg' ),
              ( ['in'], BSTR, 'unitsArg' ),
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeOutputResultAs)), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeWorksheet implementation
##class IMathcadPrimeWorksheet_Impl(object):
##    @property
##    def Name(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def FullName(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def IsReadOnly(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def _get(self):
##        '-no docstring-'
##        #return pRetVal
##    def _set(self, pRetVal):
##        '-no docstring-'
##    Modified = property(_get, _set, doc = _set.__doc__)
##
##    def SetTitle(self, titleArg):
##        '-no docstring-'
##        #return 
##
##    def Save(self):
##        '-no docstring-'
##        #return 
##
##    def SaveAs(self, newDocumentPathArg):
##        '-no docstring-'
##        #return 
##
##    def Synchronize(self):
##        '-no docstring-'
##        #return 
##
##    def PauseCalculation(self):
##        '-no docstring-'
##        #return 
##
##    def ResumeCalculation(self):
##        '-no docstring-'
##        #return 
##
##    def SetRealValue(self, aliasArg, valueArg, unitsArg):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Inputs(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Outputs(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def InputGetRealValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetRealValue(self, aliasArg):
##        '-no docstring-'
##        #return pRetVal
##
##    def OutputGetRealValueAs(self, aliasArg, unitsArg):
##        '-no docstring-'
##        #return pRetVal
##

IMathcadPrimeInputs._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Count',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2)], HRESULT, 'GetAliasByIndex',
              ( ['in'], c_int, 'indexArg' ),
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeInputs implementation
##class IMathcadPrimeInputs_Impl(object):
##    @property
##    def Count(self):
##        '-no docstring-'
##        #return pRetVal
##
##    def GetAliasByIndex(self, indexArg):
##        '-no docstring-'
##        #return pRetVal
##

IMathcadPrimeOutputMatrixResult._methods_ = [
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'ErrorCode',
              ( ['out', 'retval'], POINTER(c_int), 'pRetVal' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'MatrixResult',
              ( ['out', 'retval'], POINTER(POINTER(IMathcadPrimeMatrix)), 'pRetVal' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'Units',
              ( ['out', 'retval'], POINTER(BSTR), 'pRetVal' )),
]
################################################################
## code template for IMathcadPrimeOutputMatrixResult implementation
##class IMathcadPrimeOutputMatrixResult_Impl(object):
##    @property
##    def ErrorCode(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def MatrixResult(self):
##        '-no docstring-'
##        #return pRetVal
##
##    @property
##    def Units(self):
##        '-no docstring-'
##        #return pRetVal
##

__all__ = [ 'ValueResultTypes_Real', 'ValueResultTypes',
           'SaveOption_spPromptToSaveChanges',
           'WorksheetReadonlyOptionNames_OperationsWithEnabledStateGeneration',
           'SaveOption_spSaveChanges',
           'MathcadPrimeEvents_OnWorksheetModified',
           'MathcadPrimeEvents_OnWorksheetClosed',
           'MathcadPrimeEvents',
           'WorksheetReadonlyOptionNames_FileLocationHistoryDisabled',
           'IMathcadPrimeInputsOutputsConflicts', 'GetValueResult',
           'Worksheet', 'IMathcadPrimeMatrix',
           'WorksheetReadonlyOptions', 'ValueResultTypes_None',
           'ApplicationObsolete', 'Inputs',
           'WorksheetReadonlyOptionNames', 'ValueResultTypes_Matrix',
           'IMathcadPrimeOutputMatrixResult',
           'IMathcadPrimeOutputMatrixResultAs', 'OutputResult',
           'ValuesSetter', 'MathcadPrimeEvents_OnWorksheetRenamed',
           'InputMatrixResult',
           'MathcadPrimeEvents_OnWorksheetStatesGenerating',
           'OutputResultAs', 'InputsOutputsConflicts',
           'IMathcadPrimeSetValueResults',
           'IMathcadPrimeApplication2', 'WorksheetOperations_Save',
           'IMathcadPrimeGetValueResult',
           'MathcadPrimeEvents_OnWorksheetInputsOutputsSelected',
           'IMathcadPrimeEvents2', 'IMathcadPrimeValuesSetter',
           'Worksheets', 'InputResult', 'IMathcadPrimeApplication',
           'SaveOption', 'Matrix', 'OutputMatrixResult',
           'MathcadPrimeEvents_OnRequestToUpdateInputs',
           'MathcadPrimeEvents_OnWorksheetStatesGenerated',
           'IMathcadPrimeWorksheetReadonlyOptions',
           'WorksheetReadonlyOptionNames_CaseInsensitiveAliasComparisonEnabled',
           'IMathcadPrimeWorksheet3', 'IMathcadPrimeInputs',
           'InputsOutputsStates', 'IMathcadPrimeInputResult',
           'IMathcadPrimeWorksheet', 'IMathcadPrimeOutputs',
           'WorksheetReadonlyOptionNames_RequestToUpdateInputsEnabled',
           'IMathcadPrimeOutputResult', 'IMathcadPrimeEvents',
           'SaveOption_spDiscardChanges',
           'IMathcadPrimeOutputResultAs', 'SetValueResults',
           'ValueResultTypes_String', 'Outputs',
           'IMathcadPrimeInputMatrixResult',
           'IMathcadPrimeWorksheets',
           'IMathcadPrimeInputsOutputsStates',
           'IMathcadPrimeWorksheet2', 'MathcadPrimeEvents_OnExit',
           'MathcadPrimeEvents_OnWorksheetSaved',
           'OutputMatrixResultAs', 'IMathcadPrimeApplication3',
           'WorksheetOperations', 'Application',
           'WorksheetOperations_None']
from comtypes import _check_version; _check_version('')
