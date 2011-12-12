"""
Functions to export VBA components from Excel
workbooks.

"""
import os
import sys
import logging

from win32com import client

from vbide import COMPONENT_EXTENSION_MAP

log = logging.getLogger(__name__)

def export_vba_components(workbook_path, destination):
    if not os.path.exists(destination):
        raise StandardError('Destination %s doesn\'t exist.', destination)

    excel = client.Dispatch('Excel.Application')
    excel.visible = True
    excel.Workbooks.Open(workbook_path)
    workbook = excel.Workbooks.Item(os.path.basename(workbook_path))

    for component in workbook.VBProject.VBComponents:
        if component.Type in COMPONENT_EXTENSION_MAP:
            target = os.path.join(destination, component.Name + COMPONENT_EXTENSION_MAP[component.Type])
            logging.info('Exporting %s to %s', component.Name, target)
            component.Export(target)

    logging.info('Export of VBA component from %s to %s complete.', workbook_path, destination)
