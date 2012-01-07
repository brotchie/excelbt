#!/usr/bin/env python
"""
Dumps all macros in an given Excel workbook into
a directory as separate files.

"""
import sys
import argparse
import logging

from win32com.client import Dispatch
from excelbt import export_vba_components

if __name__ == '__main__':
    log = logging.getLogger(sys.argv[0])

    parser = argparse.ArgumentParser(description='Dumps all macros in an Excel document.')

    parser.add_argument('workbook', type=unicode, help='Workbook containing VBE macros.')
    parser.add_argument('destination', type=unicode, help='Destination folder for dumped VBE macros.')

    args = parser.parse_args()

    try:
        xl = Dispatch('Excel.Application')
        workbook = xl.Workbooks.Open(args.workbook)
        export_vba_components(workbook, args.destination)
        xl.Quit()
    except StandardError, e:
        log.error('Failed to export VBA components. Reason: %s', str(e))
        sys.exit(1)
