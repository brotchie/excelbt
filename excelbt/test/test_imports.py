from excelbt.imports import import_vba_component

def test_import_vba_component(tmpdir, xl):
    standard_module_contents = """Attribute VB_Name = "StandardModule"
' Blank standard module."""

    tmpdir.join('StandardModule.bas').write(standard_module_contents)

    xl.visible = 1
    wb = xl.Workbooks.Add()
    component = import_vba_component(wb, tmpdir.join('StandardModule.bas').strpath)
    assert component.Name == 'StandardModule'
