import py.test
py.test.importorskip('win32com')

import os
from excelbt.export import export_vba_components

DATAPATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')

def test_export(tmpdir, xl):
    expected_exports = {
        'ClassModule.cls' : ['\' Blank class module.'],
        'Sheet1.bas' : [],
        'Sheet2.bas' : [],
        'Sheet3.bas' : [],
        'StandardModule.bas' : ['\' Blank standard module.'],
        'ThisWorkbook.bas' : ['\' ThisWorkbook contents.'],
    }

    workbook_path = os.path.join(DATAPATH, 'export_test_workbook.xlsm')
    workbook = xl.Workbooks.Open(workbook_path)

    export_vba_components(workbook, tmpdir.strpath)

    basenames = set(path.basename for path in tmpdir.listdir())
    missing = set(expected_exports.keys()) - set(basenames)
    extra = set(basenames) - set(expected_exports.keys())

    assert not missing
    assert not extra

    for (filename, expected_contents) in expected_exports.iteritems():
        contents = tmpdir.join(filename).read()
        for expected_content in expected_contents:
            assert expected_content in contents
