import os
import tempfile
import shutil

def add_reference(workbook, guid, major, minor):
    workbook.VBProject.References.AddFromGuid(guid, major, minor)

def import_vba_component(workbook, source):
    """
    Imports a vba module into the given workbook.

    Workbook is a Excel Workbook COM object.

    """
    if not os.path.exists(source):
        raise StandardError('Source %s doesn\'t exist.', source)

    return workbook.VBProject.VBComponents.Import(source)

def import_vbproject(workbook, vbproject):
    """
    Imports the InMemory contents of a excelbt.vbproject.VBProject object
    into a given workbook.

    """
    components = []

    tmpdir = tempfile.mkdtemp()

    for filename in vbproject.export(tmpdir):
        print filename
        components.append(import_vba_component(workbook, os.path.join(tmpdir, filename)))

    for (guid, major, minor) in vbproject.references:
        add_reference(workbook, guid, major, minor)

    shutil.rmtree(tmpdir)

    return components
