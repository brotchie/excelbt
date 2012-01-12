import py.test
py.test.importorskip('win32com')

from excelbt.imports import import_vba_component, import_vbproject
from excelbt.vbproject import VBProject, Module, ClassModule

from win32com.client import constants as const
STANDARD_MODULE_CONTENTS = """Attribute VB_Name = "StandardModule"
' Blank standard module."""

CLASS_MODULE_CONTENTS = """VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "ClassModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
"""

def test_import_vba_component(tmpdir, xl):
    tmpdir.join('StandardModule.bas').write(STANDARD_MODULE_CONTENTS)

    xl.visible = 1
    wb = xl.Workbooks.Add()
    component = import_vba_component(wb, tmpdir.join('StandardModule.bas').strpath)
    assert component.Name == 'StandardModule'

def test_import_vbproject(tmpdir, xl):
    m1 = Module('StandardModule', STANDARD_MODULE_CONTENTS)
    m2 = ClassModule('ClassModule', CLASS_MODULE_CONTENTS)
    project = VBProject([m1, m2])

    wb = xl.Workbooks.Add()
    components = import_vbproject(wb, project)

    assert components[0].Type == const.vbext_ct_StdModule
    assert components[0].Name == 'StandardModule'

    assert components[1].Type == const.vbext_ct_ClassModule
    assert components[1].Name == 'ClassModule'
