from vs_constants import *
from _create_picture_dialog import CreatePictureDialog
from _picture import *

# import pydevd_pycharm
# pydevd_pycharm.settrace('localhost', port=12345, stdoutToServer=True, stderrToServer=True, suspend=False)


def execute() -> None:
    """ VectorWorks entry point """

    parameters = PictureParameters()
    dialog = CreatePictureDialog(parameters)
    if dialog.result == kOK:
        build_picture(parameters, None)
