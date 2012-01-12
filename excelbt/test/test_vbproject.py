from excelbt.vbproject import Module, VBProject

def test_module(tmpdir):
    m = Module('TestModule', 'Code')

    assert m.name == 'TestModule'
    assert m.code == 'Code'

    assert m.filename == 'TestModule.bas'

    m.export(tmpdir.strpath)

    assert tmpdir.join('TestModule.bas').check()
    assert tmpdir.join('TestModule.bas').read() == 'Code'

def test_vbproject(tmpdir):
    project = VBProject()

    m1 = Module('TestModule', 'Code')
    m2 = Module('ClassModule', 'Class Code')

    project.add_module(m1)
    project.add_module(m2)

    project.export(tmpdir.strpath)

    assert tmpdir.join('TestModule.bas').check()
    assert tmpdir.join('TestModule.bas').read() == 'Code'

    assert tmpdir.join('ClassModule.bas').check()
    assert tmpdir.join('ClassModule.bas').read() == 'Class Code'
