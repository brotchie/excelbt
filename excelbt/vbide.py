"""
Converts some COM specific Excel and VBA constants
into Pythonic structures.

"""
from win32com import client
from win32com.client import constants as const

client.gencache.EnsureModule('{0002E157-0000-0000-C000-000000000046}', 0, 5, 3)

SCRIPTING_REFERENCE = ('{420B2830-E718-11CF-893D-00A0C9054228}', 1, 0)

COMPONENT_EXTENSION_MAP = {
    const.vbext_ct_ClassModule : '.cls', 
    const.vbext_ct_Document : '.bas',
    const.vbext_ct_MSForm : '.frm',
    const.vbext_ct_StdModule : '.bas',
}

