'''
Created on Dec 7, 2016

@author: Carles
'''
import os
import sys
from time import gmtime, strftime
import vs
import pypyodbc as pyodbc
from vs_constants import *
from _import_settings import ImportSettings
from _import_pictures_dialog import ImportPicturesDialog


# import pydevd_pycharm
# pydevd_pycharm.settrace('localhost', port=12345, stdoutToServer=True, stderrToServer=True, suspend=False)

def execute():
    settings = ImportSettings()
    import_dialog = ImportPicturesDialog(settings)
    if import_dialog.result == kOK:
        settings.save()

# import_dialog = createImportDialog()
# if vs.RunLayoutDialog(import_dialog, importDialogHandler) == kOK:
#     settings.save()
