import os

def import_vba_component(workbook, source):
    """
    Imports a vba module into the given workbook.

    Workbook is a Excel Workbook COM object.

    """
    if not os.path.exists(source):
        raise StandardError('Source %s doesn\'t exist.', source)

    return workbook.VBProject.VBComponents.Import(source)
