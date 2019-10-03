"""
Created on Dec 7, 2016

@author: Carles
"""
# import os
# import sys
# from time import gmtime, strftime
import os
from time import strftime, gmtime
from typing import IO

import vs
from vs_constants import *
from _import_settings import ImportSettings
from _import_picture_database import ImportDatabase
from _picture_settings import PictureParameters
from _picture import build_picture

# import pypyodbc as pyodbc
# from _create_picture import imageTexture
# import pydevd
# pydevd.settrace(suspend=False)

import pydevd_pycharm

pydevd_pycharm.settrace('localhost', port=12345, stdoutToServer=True, stderrToServer=True, suspend=False)

class ImportPicturesDialog:
    def __init__(self, settings: ImportSettings):

        self.excel = ImportDatabase(settings)

        ####################################################################################
        # Widget IDs
        ####################################################################################
        self.kWidgetID_excelFileGroup = 10
        self.kWidgetID_fileNameLabel = 11
        self.kWidgetID_fileName = 12
        self.kWidgetID_fileBrowseButton = 13
        self.kWidgetID_excelSheetGroup = 14
        self.kWidgetID_excelSheetNameLabel = 15
        self.kWidgetID_excelSheetName = 16

        # Picture Image
        self.kWidgetID_imageGroup = 20
        self.kWidgetID_withImageLabel = 21
        self.kWidgetID_withImageSelector = 22
        self.kWidgetID_withImage = 23
        self.kWidgetID_imageFolderNameLabel = 24
        # self.kWidgetID_imageFolderName = 25
        # self.kWidgetID_imageFolderBrowseButton = 26
        self.kWidgetID_imageTextureLabel = 27
        self.kWidgetID_imageTextureSelector = 28
        self.kWidgetID_imageWidthLabel = 29
        self.kWidgetID_imageWidthSelector = 30
        self.kWidgetID_imageHeightLabel = 31
        self.kWidgetID_imageHeightSelector = 32
        self.kWidgetID_imagePositionLabel = 33
        self.kWidgetID_imagePositionSelector = 34
        self.kWidgetID_imagePosition = 35

        # Picture Frame
        self.kWidgetID_frameGroup = 40
        self.kWidgetID_withFrameLabel = 41
        self.kWidgetID_withFrameSelector = 42
        self.kWidgetID_withFrame = 43
        self.kWidgetID_frameWidthLabel = 44
        self.kWidgetID_frameWidthSelector = 45
        self.kWidgetID_frameHeightLabel = 46
        self.kWidgetID_frameHeightSelector = 47
        self.kWidgetID_frameThicknessLabel = 48
        self.kWidgetID_frameThicknessSelector = 49
        self.kWidgetID_frameThickness = 50
        self.kWidgetID_frameDepthLabel = 51
        self.kWidgetID_frameDepthSelector = 52
        self.kWidgetID_frameDepth = 53
        self.kWidgetID_frameClassLabel = 54
        self.kWidgetID_frameClassSelector = 55
        self.kWidgetID_frameClass = 56
        self.kWidgetID_frameTextureScaleLabel = 57
        self.kWidgetID_frameTextureScaleSelector = 58
        self.kWidgetID_frameTextureScale = 59
        self.kWidgetID_frameTextureRotationLabel = 60
        self.kWidgetID_frameTextureRotationSelector = 61
        self.kWidgetID_frameTextureRotation = 62

        # Picture Matboard
        self.kWidgetID_matboardGroup = 70
        self.kWidgetID_withMatboardLabel = 71
        self.kWidgetID_withMatboardSelector = 72
        self.kWidgetID_withMatboard = 73
        self.kWidgetID_matboardPositionLabel = 74
        self.kWidgetID_matboardPositionSelector = 75
        self.kWidgetID_matboardPosition = 76
        self.kWidgetID_matboardClassLabel = 77
        self.kWidgetID_matboardClassSelector = 78
        self.kWidgetID_matboardClass = 79
        self.kWidgetID_matboardTextureScaleLabel = 80
        self.kWidgetID_matboardTextureScaleSelector = 81
        self.kWidgetID_matboardTextureScale = 82
        self.kWidgetID_matboardTextureRotatLabel = 83
        self.kWidgetID_matboardTextureRotatSelector = 84
        self.kWidgetID_matboardTextureRotat = 85

        # Picture Glass
        self.kWidgetID_glassGroup = 90
        self.kWidgetID_withGlassLabel = 91
        self.kWidgetID_withGlassSelector = 92
        self.kWidgetID_withGlass = 93
        self.kWidgetID_glassPositionLabel = 94
        self.kWidgetID_glassPositionSelector = 95
        self.kWidgetID_glassPosition = 96
        self.kWidgetID_glassClassLabel = 97
        self.kWidgetID_glassClassSelector = 98
        self.kWidgetID_glassClass = 99

        # Import Criteria
        self.kWidgetID_excelCriteriaGroup = 100
        self.kWidgetID_excelCriteriaLabel = 101
        self.kWidgetID_excelCriteriaSelector = 102
        self.kWidgetID_excelCriteriaValue = 103

        # Create Symbol
        self.kWidgetID_SymbolGroup = 200
        self.kWidgetID_SymbolCreateSymbol = 201
        self.kWidgetID_SymbolFolderLabel = 202
        self.kWidgetID_SymbolFolderSelector = 203
        self.kWidgetID_SymbolFolder = 204

        # Import Operation
        self.kWidgetID_importGroup = 300
        self.kWidgetID_importIgnoreErrors = 301
        self.kWidgetID_importIgnoreExisting = 302
        self.kWidgetID_importIgnoreUnmodified = 303
        self.kWidgetID_importButton = 304
        self.kWidgetID_importNewCount = 305
        self.kWidgetID_importUpdatedCount = 306
        self.kWidgetID_importDeletedCount = 307
        self.kWidgetID_importErrorCount = 308
        self.kWidgetID_createMissingClasses = 309

        ####################################################################################
        # Dialog Parameters
        ####################################################################################
        self.parameters = settings

        ####################################################################################
        # Dialog Variables
        ####################################################################################
        self.importNewCount = 0
        self.importUpdatedCount = 0
        self.importDeletedCount = 0
        self.importErrorCount = 0

        # Run the dialog
        ################################################################################################################
        self.dialog = vs.CreateLayout("Import Pictures", True, "OK", "Cancel")
        self.dialog_layout()
        self.result = vs.RunLayoutDialog(self.dialog, self.dialog_handler_cb)

    def set_workbook(self) -> None:
        """ Sets a new workbook

        Thw file name of the workbook is contained in the `self.settings` object
        """
        if not self.excel.connect():
            vs.SetItemText(self.dialog, self.kWidgetID_excelSheetNameLabel, "Invalid Excel file!")
            while vs.GetChoiceCount(self.dialog, self.kWidgetID_excelSheetName):
                vs.RemoveChoice(self.dialog, self.kWidgetID_excelSheetName, 0)
            vs.ShowItem(self.dialog, self.kWidgetID_excelSheetNameLabel, True)
            vs.ShowItem(self.dialog, self.kWidgetID_excelSheetName, False)
            self.show_parameters(False)
        else:
            for worksheet in self.excel.get_worksheets():
                vs.AddChoice(self.dialog, self.kWidgetID_excelSheetName, worksheet, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_excelSheetName, "Select a worksheet", 0)

            index = vs.GetChoiceIndex(self.dialog, self.kWidgetID_excelSheetName, self.parameters.excelSheetName)
            if index == -1:
                vs.SelectChoice(self.dialog, self.kWidgetID_excelSheetName, 0, True)
                self.parameters.excelSheetName = "Select a worksheet"
            else:
                vs.SelectChoice(self.dialog, self.kWidgetID_excelSheetName, index, True)
                self.show_parameters(True)

            # vs.SetItemText(self.dialog, self.kWidgetID_excelSheetNameLabel, "Excel sheet: ")
            vs.ShowItem(self.dialog, self.kWidgetID_excelSheetNameLabel, True)
            vs.ShowItem(self.dialog, self.kWidgetID_excelSheetName, True)

    def update_criteria_values(self, state) -> None:
        """ Updates the criteria field

        :rtype: None
        """

        criteria_values = self.excel.get_criteria_values()
        if criteria_values and state is True and self.parameters.excelCriteriaSelector != "-- Select column ...":
            for criteria in criteria_values:
                vs.AddChoice(self.dialog, self.kWidgetID_excelCriteriaValue, criteria, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_excelCriteriaValue, "Select a value ...", 0)
            index = vs.GetChoiceIndex(self.dialog, self.kWidgetID_excelCriteriaValue,
                                      self.parameters.excelCriteriaValue)
            if index == -1:
                vs.SelectChoice(self.dialog, self.kWidgetID_excelCriteriaValue, 0, True)
                self.parameters.excelCriteriaValue = "Select a value ..."
            else:
                vs.SelectChoice(self.dialog, self.kWidgetID_excelCriteriaValue, index, True)
        else:
            while vs.GetChoiceCount(self.dialog, self.kWidgetID_excelCriteriaValue):
                vs.RemoveChoice(self.dialog, self.kWidgetID_excelCriteriaValue, 0)

    def dialog_handler_cb(self, item, data) -> None:
        """ Handles dialog events

        This is a call-back function invoked by VectorWorks whenever there is a change in the state od the dialog box.
        Changes in the dialog fields will be reflected in the `self.settings` object.

        :param item: The ID of the field of dialog box life cycle event
        :param data: Data associated with the event
        :returns: None

        """
        # Dialog box initialization event
        if item == KDialogInitEvent:
            vs.SetItemText(self.dialog, self.kWidgetID_fileName, self.parameters.excelFileName)
            # vs.SetItemText(self.dialog, self.kWidgetID_imageFolderName, self.settings.imageFolderName)

            vs.ShowItem(self.dialog, self.kWidgetID_excelSheetNameLabel, False)
            vs.ShowItem(self.dialog, self.kWidgetID_excelSheetName, False)
            self.show_parameters(False)

            vs.EnableItem(self.dialog, self.kWidgetID_importButton, False)
            vs.EnableItem(self.dialog, self.kWidgetID_importNewCount, False)
            vs.EnableItem(self.dialog, self.kWidgetID_importUpdatedCount, False)
            vs.EnableItem(self.dialog, self.kWidgetID_importDeletedCount, False)

        elif item == self.kWidgetID_fileName:
            self.parameters.excelFileName = vs.GetItemText(self.dialog, self.kWidgetID_fileName)

        elif item == self.kWidgetID_fileBrowseButton:
            result, self.parameters.excelFileName = vs.GetFileN("Open Excel file", "", "xlsm")
            if result:
                vs.SetItemText(self.dialog, self.kWidgetID_fileName, self.parameters.excelFileName)

        elif item == self.kWidgetID_excelSheetName:
            new_excel_sheet_name = vs.GetChoiceText(self.dialog, self.kWidgetID_excelSheetName, data)
            if self.parameters.excelSheetName != new_excel_sheet_name:
                self.parameters.excelSheetName = new_excel_sheet_name
                self.show_parameters(False)
                if data != 0:
                    self.show_parameters(True)

        elif item == self.kWidgetID_withImageSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_withImage, data == 0)
            self.parameters.withImageSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_withImageSelector, data)
        elif item == self.kWidgetID_withImage:
            self.parameters.pictureParameters.withImage = "{}".format(data != 0)
        # elif item == self.kWidgetID_imageFolderName:
        #     self.settings.imageFolderName = vs.GetItemText(
        #         self.dialog, self.kWidgetID_imageFolderName)
        # elif item == self.kWidgetID_imageFolderBrowseButton:
        #     result, self.settings.imageFolderName = vs.GetFolder("Select the images folder")
        #     if result == 0:
        #         vs.SetItemText(self.dialog, self.kWidgetID_imageFolderName, self.settings.imageFolderName)
        elif item == self.kWidgetID_imageTextureSelector:
            self.parameters.imageTextureSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_withImageSelector, data)
        elif item == self.kWidgetID_imageWidthSelector:
            self.parameters.imageWidthSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_imageWidthSelector, data)
        elif item == self.kWidgetID_imageHeightSelector:
            self.parameters.imageHeightSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_imageHeightSelector, data)
        elif item == self.kWidgetID_imagePositionSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_imagePosition, data == 0)
            self.parameters.imagePositionSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_imagePositionSelector, data)
        elif item == self.kWidgetID_imagePosition:
            _, self.parameters.pictureParameters.imagePosition = str(vs.GetEditReal(
                self.dialog, self.kWidgetID_imagePosition, 3))
        elif item == self.kWidgetID_withFrameSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_withFrame, data == 0)
            self.parameters.withFrameSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_withFrameSelector, data)
        elif item == self.kWidgetID_withFrame:
            self.parameters.pictureParameters.withFrame = "{}".format(data != 0)
        elif item == self.kWidgetID_frameWidthSelector:
            self.parameters.frameWidthSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_frameWidthSelector, data)
        elif item == self.kWidgetID_frameHeightSelector:
            self.parameters.frameHeightSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_frameHeightSelector, data)
        elif item == self.kWidgetID_frameThicknessSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_frameThickness, data == 0)
            self.parameters.frameThicknessSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_frameThicknessSelector, data)
        elif item == self.kWidgetID_frameThickness:
            _, self.parameters.pictureParameters.frameThickness = str(vs.GetEditReal(
                self.dialog, self.kWidgetID_frameThickness, 3))
        elif item == self.kWidgetID_frameDepthSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_frameDepth, data == 0)
            self.parameters.frameDepthSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_frameDepthSelector, data)
        elif item == self.kWidgetID_frameDepth:
            _, self.parameters.pictureParameters.frameDepth = str(vs.GetEditReal(
                self.dialog, self.kWidgetID_frameDepth, 3))
        elif item == self.kWidgetID_frameClassSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_frameClass, data == 0)
            self.parameters.frameClassSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_frameClassSelector, data)
        elif item == self.kWidgetID_frameClass:
            index, self.parameters.frameClass = vs.GetSelectedChoiceInfo(
                self.dialog, self.kWidgetID_frameClass, 0)
        elif item == self.kWidgetID_frameTextureScaleSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScale, data == 0)
            self.parameters.frameTextureScaleSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_frameTextureScaleSelector, data)
        elif item == self.kWidgetID_frameTextureScale:
            _, self.parameters.pictureParameters.frameTextureScale = str(vs.GetEditReal(
                self.dialog, self.kWidgetID_frameTextureScale, 1))
        elif item == self.kWidgetID_frameTextureRotationSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotation, data == 0)
            self.parameters.frameTextureRotationSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_frameTextureRotationSelector, data)
        elif item == self.kWidgetID_frameTextureRotation:
            _, self.parameters.pictureParameters.frameTextureRotation = str(vs.GetEditReal(
                self.dialog, self.kWidgetID_frameTextureRotation, 1))
        elif item == self.kWidgetID_withMatboardSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_withMatboard, data == 0)
            self.parameters.withMatboardSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_withMatboardSelector, data)
        elif item == self.kWidgetID_withMatboard:
            self.parameters.pictureParameters.withMatboard = "{}".format(data is True)
        elif item == self.kWidgetID_matboardPositionSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_matboardPosition, data == 0)
            self.parameters.matboardPositionSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_matboardPositionSelector, data)
        elif item == self.kWidgetID_matboardPosition:
            _, self.parameters.pictureParameters.matboardPosition = str(vs.GetEditReal(
                self.dialog, self.kWidgetID_matboardPosition, 3))
        elif item == self.kWidgetID_matboardClassSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_matboardClass, data == 0)
            self.parameters.matboardClassSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_matboardClassSelector, data)
        elif item == self.kWidgetID_matboardClass:
            index, self.parameters.matboardClass = vs.GetSelectedChoiceInfo(
                self.dialog, self.kWidgetID_matboardClass, 0)
        elif item == self.kWidgetID_matboardTextureScaleSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScale, data == 0)
            self.parameters.matboardTextureScaleSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_matboardTextureScaleSelector, data)
        elif item == self.kWidgetID_matboardTextureScale:
            _, self.parameters.pictureParameters.matboardTextureScale = str(vs.GetEditReal(
                self.dialog, self.kWidgetID_matboardTextureScale, 1))
        elif item == self.kWidgetID_matboardTextureRotatSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotat, data == 0)
            self.parameters.matboardTextureRotatSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_matboardTextureRotatSelector, data)
        elif item == self.kWidgetID_matboardTextureRotat:
            _, self.parameters.pictureParameters.matboardTextureRotat = str(vs.GetEditReal(
                self.dialog, self.kWidgetID_matboardTextureRotat, 1))
        elif item == self.kWidgetID_withGlassSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_withGlass, data == 0)
            self.parameters.withGlassSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_withGlassSelector, data)
        elif item == self.kWidgetID_withGlass:
            self.parameters.pictureParameters.withGlass = "{}".format(data is True)
        elif item == self.kWidgetID_glassPositionSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_glassPosition, data == 0)
            self.parameters.glassPositionSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_glassPositionSelector, data)
        elif item == self.kWidgetID_glassPosition:
            _, self.parameters.pictureParameters.glassPosition = str(vs.GetEditReal(
                self.dialog, self.kWidgetID_glassPosition, 3))
        elif item == self.kWidgetID_glassClassSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_glassClass, data == 0)
            self.parameters.glassClassSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_glassClassSelector, data)
        elif item == self.kWidgetID_glassClass:
            index, self.parameters.glassClass = vs.GetSelectedChoiceInfo(
                self.dialog, self.kWidgetID_glassClass, 0)
        elif item == self.kWidgetID_excelCriteriaSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_excelCriteriaValue, data != 0)
            new_excel_criteria_selector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_excelCriteriaSelector, data)
            if new_excel_criteria_selector != self.parameters.excelCriteriaSelector:
                self.parameters.excelCriteriaSelector = new_excel_criteria_selector
                self.update_criteria_values(False)
                if data != 0:
                    self.update_criteria_values(True)
                else:
                    index = vs.GetChoiceIndex(
                        self.dialog, self.kWidgetID_excelCriteriaValue, self.parameters.excelCriteriaValue)
                    if index == -1:
                        vs.SelectChoice(self.dialog, self.kWidgetID_excelCriteriaValue, 0, True)
                        self.parameters.excelCriteriaValue = "Select a value ..."
                    else:
                        vs.SelectChoice(self.dialog, self.kWidgetID_excelCriteriaValue, index, True)
        elif item == self.kWidgetID_excelCriteriaValue:
            self.parameters.excelCriteriaValue = vs.GetChoiceText(
                self.dialog, self.kWidgetID_excelCriteriaValue, data)
        elif item == self.kWidgetID_SymbolCreateSymbol:
            self.parameters.symbolCreateSymbol = "{}".format(data)
            state = vs.GetSelectedChoiceIndex(
                self.dialog, self.kWidgetID_SymbolFolderSelector, 0) == 0 and data is True
            vs.EnableItem(self.dialog, self.kWidgetID_SymbolFolderSelector, data)
            vs.EnableItem(self.dialog, self.kWidgetID_SymbolFolder, state)
        elif item == self.kWidgetID_SymbolFolderSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_SymbolFolder, data == 0)
            self.parameters.symbolFolderSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_SymbolFolderSelector, data)
        elif item == self.kWidgetID_importIgnoreErrors:
            self.parameters.importIgnoreErrors = "{}".format(data is True)
            vs.ShowItem(self.dialog, self.kWidgetID_importErrorCount, data is not True)
        elif item == self.kWidgetID_importIgnoreExisting:
            self.parameters.importIgnoreExisting = "{}".format(data is True)
        elif item == self.kWidgetID_importIgnoreUnmodified:
            self.parameters.importIgnoreUnmodified = "{}".format(data is True)
        elif item == self.kWidgetID_createMissingClasses:
            self.parameters.createMissingClasses = "{}".format(data is True)
        elif item == self.kWidgetID_importButton:
            self.import_pictures()
            vs.SetItemText(self.dialog, self.kWidgetID_importNewCount,
                           "New Pictures: {}".format(self.importNewCount))
            vs.SetItemText(self.dialog, self.kWidgetID_importUpdatedCount,
                           "Updated Pictures: {}".format(self.importUpdatedCount))
            vs.SetItemText(self.dialog, self.kWidgetID_importDeletedCount,
                           "Deleted Pictures: {}".format(self.importDeletedCount))
            vs.SetItemText(self.dialog, self.kWidgetID_importErrorCount,
                           "Error Pictures: {}".format(self.importErrorCount))

        # This section handles the following cases:
        # - The Dialog is initializing
        # - The name of the workbook file has changed
        if item == self.kWidgetID_fileName or \
                item == self.kWidgetID_fileBrowseButton or \
                item == KDialogInitEvent:
            self.set_workbook()

        # The image selection has changed
        if item == self.kWidgetID_withImageSelector or \
                item == self.kWidgetID_withImage or \
                item == self.kWidgetID_excelSheetName:
            state = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_withImageSelector, 0) != 0 or \
                    vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage) is True

            vs.EnableItem(self.dialog, self.kWidgetID_imageWidthLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_imageWidthSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_imageHeightLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_imageHeightSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_imagePositionLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_imagePositionSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_imagePosition, state)
            vs.EnableItem(self.dialog, self.kWidgetID_imageTextureLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_imageTextureSelector, state)

        # The frame selection has changed
        if item == self.kWidgetID_withFrameSelector or \
                item == self.kWidgetID_withFrame or \
                item == self.kWidgetID_excelSheetName:
            state = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_withFrameSelector, 0) != 0 or \
                    vs.GetBooleanItem(self.dialog, self.kWidgetID_withFrame) is True

            vs.EnableItem(self.dialog, self.kWidgetID_frameWidthLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameWidthSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameHeightLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameHeightSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameThicknessLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameThicknessSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameThickness, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameDepthLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameDepthSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameDepth, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameClassLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameClassSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameClass, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScaleLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScaleSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScale, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotationLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotationSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotation, state)

        # The matboard selection has changed
        if item == self.kWidgetID_withMatboardSelector or \
                item == self.kWidgetID_withMatboard or \
                item == self.kWidgetID_excelSheetName:
            state = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_withMatboardSelector, 0) != 0 or \
                    vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard) is True

            vs.EnableItem(self.dialog, self.kWidgetID_matboardPositionLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardPositionSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardPosition, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardClassLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardClassSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardClass, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScaleLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScaleSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScale, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotatLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotatSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotat, state)

        # The glass selection has changed
        if item == self.kWidgetID_withGlassSelector or \
                item == self.kWidgetID_withGlass or \
                item == self.kWidgetID_excelSheetName:
            state = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_withGlassSelector, 0) != 0 or \
                    vs.GetBooleanItem(self.dialog, self.kWidgetID_withGlass) is True

            vs.EnableItem(self.dialog, self.kWidgetID_glassPositionLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_glassPositionSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_glassPosition, state)
            vs.EnableItem(self.dialog, self.kWidgetID_glassClassLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_glassClassSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_glassClass, state)

        # After the event has been handled, update some of the import validity settings accordingly
        self.parameters.imageValid = \
            ((self.parameters.withImageSelector == "-- Manual" and
              self.parameters.pictureParameters.withImage == "True") or
             self.parameters.withImageSelector != "-- Manual") and \
            (self.parameters.imageTextureSelector != "-- Select column ...") and \
            (self.parameters.imageWidthSelector != "-- Select column ...") and \
            (self.parameters.imageHeightSelector != "-- Select column ...")

        self.parameters.frameValid = \
            ((self.parameters.withFrameSelector == "-- Manual" and
              self.parameters.pictureParameters.withFrame == "True") or
             self.parameters.withFrameSelector != "-- Manual") and \
            (self.parameters.frameWidthSelector != "-- Select column ...") and \
            (self.parameters.frameHeightSelector != "-- Select column ...")

        self.parameters.matboardValid = \
            ((self.parameters.withMatboardSelector == "-- Manual" and
              self.parameters.pictureParameters.withMatboard == "True") or
             self.parameters.withMatboardSelector != "-- Manual")

        self.parameters.glassValid = \
            ((self.parameters.withGlassSelector == "-- Manual" and
              self.parameters.pictureParameters.withGlass == "True") or
             self.parameters.withGlassSelector != "-- Manual")

        self.parameters.criteriaValid = \
            (self.parameters.excelCriteriaSelector != "-- Select column ..." and
             self.parameters.excelCriteriaValue != "Select a value ...")

        self.parameters.importValid = \
            (self.parameters.imageValid or self.parameters.frameValid) and self.parameters.criteriaValid

        vs.EnableItem(self.dialog, self.kWidgetID_importButton, self.parameters.importValid)
        vs.EnableItem(self.dialog, self.kWidgetID_importNewCount, self.parameters.importValid)
        vs.EnableItem(self.dialog, self.kWidgetID_importUpdatedCount, self.parameters.importValid)
        vs.EnableItem(self.dialog, self.kWidgetID_importDeletedCount, self.parameters.importValid)

    def show_parameters(self, state):

        vs.ShowItem(self.dialog, self.kWidgetID_imageGroup, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withImageLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withImageSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withImage, state)
        vs.ShowItem(self.dialog, self.kWidgetID_imageFolderNameLabel, state)
        # vs.ShowItem(self.dialog, self.kWidgetID_imageFolderName, state)
        # vs.ShowItem(self.dialog, self.kWidgetID_imageFolderBrowseButton, state)
        vs.ShowItem(self.dialog, self.kWidgetID_imageWidthLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_imageWidthSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_imageHeightLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_imageHeightSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_imagePositionLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_imagePositionSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_imagePosition, state)
        vs.ShowItem(self.dialog, self.kWidgetID_imageTextureLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_imageTextureSelector, state)

        vs.ShowItem(self.dialog, self.kWidgetID_frameGroup, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withFrameLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withFrameSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withFrame, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameWidthLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameWidthSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameHeightLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameHeightSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameThicknessLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameThicknessSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameThickness, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameDepthLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameDepthSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameDepth, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameClassLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameClassSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameClass, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameTextureScaleLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameTextureScaleSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameTextureScale, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameTextureRotationLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameTextureRotationSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_frameTextureRotation, state)

        vs.ShowItem(self.dialog, self.kWidgetID_matboardGroup, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withMatboardLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withMatboardSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withMatboard, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardPositionLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardPositionSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardPosition, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardClassLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardClassSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardClass, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardTextureScaleLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardTextureScaleSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardTextureScale, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardTextureRotatLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardTextureRotatSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_matboardTextureRotat, state)

        vs.ShowItem(self.dialog, self.kWidgetID_glassGroup, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withGlassLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withGlassSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_withGlass, state)
        vs.ShowItem(self.dialog, self.kWidgetID_glassPositionLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_glassPositionSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_glassPosition, state)
        vs.ShowItem(self.dialog, self.kWidgetID_glassClassLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_glassClassSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_glassClass, state)

        vs.ShowItem(self.dialog, self.kWidgetID_excelCriteriaGroup, state)
        vs.ShowItem(self.dialog, self.kWidgetID_excelCriteriaLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_excelCriteriaSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_excelCriteriaValue, state)

        vs.ShowItem(self.dialog, self.kWidgetID_SymbolGroup, state)
        vs.ShowItem(self.dialog, self.kWidgetID_SymbolCreateSymbol, state)
        vs.ShowItem(self.dialog, self.kWidgetID_SymbolFolderLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_SymbolFolderSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_SymbolFolder, state)

        vs.ShowItem(self.dialog, self.kWidgetID_importGroup, state)
        vs.ShowItem(self.dialog, self.kWidgetID_createMissingClasses, state)
        vs.ShowItem(self.dialog, self.kWidgetID_importIgnoreErrors, state)
        vs.ShowItem(self.dialog, self.kWidgetID_importIgnoreExisting, state)
        vs.ShowItem(self.dialog, self.kWidgetID_importButton, state)
        vs.ShowItem(self.dialog, self.kWidgetID_importNewCount, state)
        vs.ShowItem(self.dialog, self.kWidgetID_importUpdatedCount, state)
        vs.ShowItem(self.dialog, self.kWidgetID_importDeletedCount, state)
        vs.ShowItem(self.dialog, self.kWidgetID_importErrorCount,
                    state and self.parameters.importIgnoreErrors != "True")

        if state is False:
            self.empty_all_fields()
            self.update_criteria_values(False)

        else:
            self.populate_all_fields()

    def populate_all_fields(self):
        columns = self.excel.get_columns()
        for column in columns:
            vs.AddChoice(self.dialog, self.kWidgetID_withImageSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_imageWidthSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_imageHeightSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_imagePositionSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_imageTextureSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_withFrameSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_frameWidthSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_frameHeightSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_frameThicknessSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_frameDepthSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_frameClassSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_frameTextureScaleSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_frameTextureRotationSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_withMatboardSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_matboardPositionSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_matboardClassSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_matboardTextureScaleSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_matboardTextureRotatSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_withGlassSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_glassPositionSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_glassClassSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_excelCriteriaSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_SymbolFolderSelector, column, 0)

        vs.AddChoice(self.dialog, self.kWidgetID_withImageSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_imageTextureSelector, "-- Select column ...", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_imageWidthSelector, "-- Select column ...", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_imageHeightSelector, "-- Select column ...", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_imagePositionSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_withFrameSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_frameWidthSelector, "-- Select column ...", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_frameHeightSelector, "-- Select column ...", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_frameThicknessSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_frameDepthSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_frameClassSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_frameTextureScaleSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_frameTextureRotationSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_withMatboardSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_matboardPositionSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_matboardClassSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_matboardTextureScaleSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_matboardTextureRotatSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_withGlassSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_glassPositionSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_glassClassSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_excelCriteriaSelector, "-- Select column ...", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_SymbolFolderSelector, "-- Manual", 0)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_withImageSelector, self.parameters.withImageSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_withImageSelector, selector_index, True)

        vs.SetBooleanItem(self.dialog, self.kWidgetID_withImage, self.parameters.pictureParameters.withImage == "True")

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_imageTextureSelector,
                                                self.parameters.imageTextureSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_imageTextureSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_imageWidthSelector,
                                                self.parameters.imageWidthSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_imageWidthSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_imageHeightSelector,
                                                self.parameters.imageHeightSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_imageHeightSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_imagePositionSelector,
                                                self.parameters.imagePositionSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_imagePositionSelector, selector_index, True)

        #            vs.SetItemText(self.dialog, self.kWidgetID_imagePosition, imagePosition)
        vs.SetEditReal(self.dialog, self.kWidgetID_imagePosition, 3, self.parameters.pictureParameters.imagePosition)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_withFrameSelector,
                                                self.parameters.withFrameSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_withFrameSelector, selector_index, True)

        vs.SetBooleanItem(self.dialog, self.kWidgetID_withFrame, self.parameters.pictureParameters.withFrame == "True")

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_frameWidthSelector,
                                                self.parameters.frameWidthSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameWidthSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_frameHeightSelector,
                                                self.parameters.frameHeightSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameHeightSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_frameThicknessSelector,
                                                self.parameters.frameThicknessSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameThicknessSelector, selector_index, True)

        #            vs.SetItemText(self.dialog, self.kWidgetID_frameThickness, frameThickness)
        vs.SetEditReal(self.dialog, self.kWidgetID_frameThickness, 3, self.parameters.pictureParameters.frameThickness)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_frameDepthSelector,
                                                self.parameters.frameDepthSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameDepthSelector, selector_index, True)

        #            vs.SetItemText(self.dialog, self.kWidgetID_frameDepth, frameDepth)
        vs.SetEditReal(self.dialog, self.kWidgetID_frameDepth, 3, self.parameters.pictureParameters.frameDepth)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_frameClassSelector,
                                                self.parameters.frameClassSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameClassSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_frameClass,
                                                self.parameters.pictureParameters.frameClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameClass, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_frameTextureScaleSelector,
                                                self.parameters.frameTextureScaleSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameTextureScaleSelector, selector_index, True)

        #            vs.SetItemText(self.dialog, self.kWidgetID_frameTextureScale, frameTextureScale)
        vs.SetEditReal(self.dialog, self.kWidgetID_frameTextureScale, 1,
                       self.parameters.pictureParameters.frameTextureScale)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_frameTextureRotationSelector,
                                                self.parameters.frameTextureRotationSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameTextureRotationSelector, selector_index, True)

        #            vs.SetItemText(self.dialog, self.kWidgetID_frameTextureRotation, frameTextureRotation)
        vs.SetEditReal(self.dialog, self.kWidgetID_frameTextureRotation, 1,
                       self.parameters.pictureParameters.frameTextureRotation)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_withMatboardSelector,
                                                self.parameters.withMatboardSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_withMatboardSelector, selector_index, True)

        vs.SetBooleanItem(self.dialog, self.kWidgetID_withMatboard,
                          self.parameters.pictureParameters.withMatboard == "True")

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_matboardPositionSelector,
                                                self.parameters.matboardPositionSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardPositionSelector, selector_index, True)

        #            vs.SetItemText(self.dialog, self.kWidgetID_matboardPosition, matboardPosition)
        vs.SetEditReal(self.dialog, self.kWidgetID_matboardPosition, 3,
                       self.parameters.pictureParameters.matboardPosition)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_matboardClassSelector,
                                                self.parameters.matboardClassSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardClassSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_matboardClass,
                                                self.parameters.pictureParameters.matboardClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardClass, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_matboardTextureScaleSelector,
                                                self.parameters.matboardTextureScaleSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardTextureScaleSelector, selector_index, True)

        #            vs.SetItemText(self.dialog, self.kWidgetID_matboardTextureScale, matboardTextureScale)
        vs.SetEditReal(self.dialog, self.kWidgetID_matboardTextureScale, 1,
                       self.parameters.pictureParameters.matboardTextureScale)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_matboardTextureRotatSelector,
                                                self.parameters.matboardTextureRotatSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardTextureRotatSelector, selector_index, True)

        #            vs.SetItemText(self.dialog, self.kWidgetID_matboardTextureRotat, matboardTextureRotat)
        vs.SetEditReal(self.dialog, self.kWidgetID_matboardTextureRotat, 1,
                       self.parameters.pictureParameters.matboardTextureRotat)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_withGlassSelector,
                                                self.parameters.withGlassSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_withGlassSelector, selector_index, True)

        vs.SetBooleanItem(self.dialog, self.kWidgetID_withGlass,
                          self.parameters.pictureParameters.withGlass == "True")

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_glassPositionSelector,
                                                self.parameters.glassPositionSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_glassPositionSelector, selector_index, True)

        #            vs.SetItemText(self.dialog, self.kWidgetID_glassPosition, glassPosition)
        vs.SetEditReal(self.dialog, self.kWidgetID_glassPosition, 3,
                       self.parameters.pictureParameters.glassPosition)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_glassClassSelector,
                                                self.parameters.glassClassSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_glassClassSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_glassClass,
                                                self.parameters.pictureParameters.glassClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_glassClass, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_excelCriteriaSelector,
                                                self.parameters.excelCriteriaSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_excelCriteriaSelector, selector_index, True)

        vs.SetBooleanItem(self.dialog, self.kWidgetID_SymbolCreateSymbol,
                          self.parameters.symbolCreateSymbol == "True")

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_SymbolFolderSelector,
                                                self.parameters.symbolFolderSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_SymbolFolderSelector, selector_index, True)

        self.update_criteria_values(True)

        vs.EnableItem(self.dialog, self.kWidgetID_withImage,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_withImageSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_imagePosition,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_imagePositionSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_withFrame,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_withFrameSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_frameThickness,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_frameThicknessSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_frameDepth,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_frameDepthSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_frameClass,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_frameClassSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScale,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_frameTextureScaleSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotation,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_frameTextureRotationSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_withMatboard,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_withMatboardSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardPosition,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_matboardPositionSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardClass,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_matboardClassSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScale,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_matboardTextureScaleSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotat,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_matboardTextureRotatSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_withGlass,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_withGlassSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_glassPosition,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_glassPositionSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_glassClass,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_glassClassSelector, 0) == 0)
        vs.EnableItem(self.dialog, self.kWidgetID_excelCriteriaValue,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_excelCriteriaSelector, 0) != 0)
        vs.EnableItem(self.dialog, self.kWidgetID_SymbolFolder,
                      vs.GetSelectedChoiceIndex(self.dialog,
                                                self.kWidgetID_SymbolFolderSelector, 0) == 0)

        vs.SetBooleanItem(self.dialog, self.kWidgetID_importIgnoreErrors,
                          self.parameters.importIgnoreErrors == "True")
        vs.SetBooleanItem(self.dialog, self.kWidgetID_importIgnoreExisting,
                          self.parameters.importIgnoreExisting == "True")
        vs.SetBooleanItem(self.dialog, self.kWidgetID_importIgnoreUnmodified,
                          self.parameters.importIgnoreUnmodified == "True")
        vs.SetBooleanItem(self.dialog, self.kWidgetID_createMissingClasses,
                          self.parameters.createMissingClasses == "True")

    def remove_field_options(self, widget_id: int):
        while vs.GetChoiceCount(self.dialog, widget_id):
            vs.RemoveChoice(self.dialog, widget_id, 0)

    def empty_all_fields(self):
        self.remove_field_options(self.kWidgetID_withImageSelector)
        self.remove_field_options(self.kWidgetID_imageTextureSelector)
        self.remove_field_options(self.kWidgetID_imageWidthSelector)
        self.remove_field_options(self.kWidgetID_imageHeightSelector)
        self.remove_field_options(self.kWidgetID_imagePositionSelector)
        self.remove_field_options(self.kWidgetID_withFrameSelector)
        self.remove_field_options(self.kWidgetID_frameWidthSelector)
        self.remove_field_options(self.kWidgetID_frameHeightSelector)
        self.remove_field_options(self.kWidgetID_frameThicknessSelector)
        self.remove_field_options(self.kWidgetID_frameDepthSelector)
        self.remove_field_options(self.kWidgetID_frameClassSelector)
        self.remove_field_options(self.kWidgetID_frameTextureScaleSelector)
        self.remove_field_options(self.kWidgetID_frameTextureRotationSelector)
        self.remove_field_options(self.kWidgetID_withMatboardSelector)
        self.remove_field_options(self.kWidgetID_matboardPositionSelector)
        self.remove_field_options(self.kWidgetID_matboardClassSelector)
        self.remove_field_options(self.kWidgetID_matboardTextureScaleSelector)
        self.remove_field_options(self.kWidgetID_matboardTextureRotatSelector)
        self.remove_field_options(self.kWidgetID_withGlassSelector)
        self.remove_field_options(self.kWidgetID_glassPositionSelector)
        self.remove_field_options(self.kWidgetID_glassClassSelector)
        self.remove_field_options(self.kWidgetID_excelCriteriaSelector)
        self.remove_field_options(self.kWidgetID_SymbolFolderSelector)

    def dialog_layout(self):

        input_field_width = 15
        label_width = 20

        # Excel file group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_excelFileGroup, "Excel spreadsheet", True)
        vs.SetFirstLayoutItem(self.dialog, self.kWidgetID_excelFileGroup)
        # File Name
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_fileNameLabel, "Excel file: ", -1)
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_excelFileGroup, self.kWidgetID_fileNameLabel)
        vs.CreateEditText(self.dialog, self.kWidgetID_fileName, self.parameters.excelFileName, 3 * input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_fileNameLabel, self.kWidgetID_fileName, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_fileName, "Enter the excel file name here")
        # File browse button
        # -----------------------------------------------------------------------------------------
        vs.CreatePushButton(self.dialog, self.kWidgetID_fileBrowseButton, "Browse...")
        vs.SetRightItem(self.dialog, self.kWidgetID_fileName, self.kWidgetID_fileBrowseButton, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_fileBrowseButton, "Click to browse Excel file")
        # Excel sheet selection
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_excelSheetNameLabel, "Excel sheet: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_fileNameLabel, self.kWidgetID_excelSheetNameLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_excelSheetName, input_field_width)
        sheet_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_excelSheetName, self.parameters.excelSheetName)
        vs.SelectChoice(self.dialog, self.kWidgetID_excelSheetName, sheet_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_excelSheetNameLabel, self.kWidgetID_excelSheetName, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_excelSheetName, "Select the Excel sheet")

        # Image group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_imageGroup, "Image", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_excelFileGroup, self.kWidgetID_imageGroup, 0, 0)
        # With Image checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_withImageLabel, "With Image: ", label_width)
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_imageGroup, self.kWidgetID_withImageLabel)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_withImageSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_withImageSelector, self.parameters.withImageSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_withImageSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_withImageLabel, self.kWidgetID_withImageSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_withImageSelector, "Select the column for the image creation")
        vs.CreateCheckBox(self.dialog, self.kWidgetID_withImage, "Include Image   ")
        vs.SetBooleanItem(self.dialog, self.kWidgetID_withImage, self.parameters.pictureParameters.withImage == "True")
        vs.SetRightItem(self.dialog, self.kWidgetID_withImageSelector, self.kWidgetID_withImage, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_withImage, "Choose the value for the image creation")
        # Image Folder Name
        # -----------------------------------------------------------------------------------------
        # vs.CreateStaticText(self.dialog, self.kWidgetID_imageFolderNameLabel, "Images folder: ", label_width)
        # vs.SetBelowItem(self.dialog, self.kWidgetID_withImageLabel, self.kWidgetID_imageFolderNameLabel, 0, 0)
        # vs.CreateEditText(self.dialog, self.kWidgetID_imageFolderName,
        #                   self.settings.imageFolderName, input_field_width)
        # vs.SetRightItem(self.dialog, self.kWidgetID_imageFolderNameLabel, self.kWidgetID_imageFolderName, 0, 0)
        # vs.SetHelpText(self.dialog, self.kWidgetID_imageFolderName, "Enter the folder for the image files")
        # File browse button
        # -----------------------------------------------------------------------------------------
        # vs.CreatePushButton(self.dialog, self.kWidgetID_imageFolderBrowseButton, "Browse...")
        # vs.SetRightItem(self.dialog, self.kWidgetID_imageFolderName, self.kWidgetID_imageFolderBrowseButton, 0, 0)
        # vs.SetHelpText(self.dialog, self.kWidgetID_imageFolderBrowseButton, "Click to browse the images folder")
        # Image Texture
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_imageTextureLabel, "Image name: ", label_width)
        # vs.SetBelowItem(self.dialog, self.kWidgetID_imageFolderNameLabel, self.kWidgetID_imageTextureLabel, 0, 0)
        vs.SetBelowItem(self.dialog, self.kWidgetID_withImageLabel, self.kWidgetID_imageTextureLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_imageTextureSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_imageTextureSelector,
                                                self.parameters.imageTextureSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_imageTextureSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_imageTextureLabel, self.kWidgetID_imageTextureSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_imageTextureSelector, "Select the column for the image name")
        # Image Width dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_imageWidthLabel, "Image Width: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_imageTextureLabel, self.kWidgetID_imageWidthLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_imageWidthSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_imageWidthSelector,
                                                self.parameters.imageWidthSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_imageWidthSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_imageWidthLabel, self.kWidgetID_imageWidthSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_imageWidthSelector, "Select the column for the image width")
        # Image Height dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_imageHeightLabel, "Image Height: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_imageWidthLabel, self.kWidgetID_imageHeightLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_imageHeightSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_imageHeightSelector,
                                                self.parameters.imageHeightSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_imageHeightSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_imageHeightLabel, self.kWidgetID_imageHeightSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_imageHeightSelector, "Select the column for the image height")
        # Image Position dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_imagePositionLabel, "Image Position: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_imageHeightLabel, self.kWidgetID_imagePositionLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_imagePositionSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_imagePositionSelector,
                                                self.parameters.imagePositionSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_imagePositionSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_imagePositionLabel, self.kWidgetID_imagePositionSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_imagePositionSelector, "Select the column for the image position")
        valid, value = vs.ValidNumStr(self.parameters.pictureParameters.imagePosition)
        if not valid:
            value = PictureParameters().imagePosition
        vs.CreateEditReal(self.dialog, self.kWidgetID_imagePosition,
                          3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_imagePositionSelector, self.kWidgetID_imagePosition, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_imagePosition, "Enter the position (depth) of the image here.")

        # Frame group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_frameGroup, "Frame", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_imageGroup, self.kWidgetID_frameGroup, 0, 0)
        # With Frame checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_withFrameLabel, "With Frame: ", label_width)
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_frameGroup, self.kWidgetID_withFrameLabel)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_withFrameSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_withFrameSelector,
                                                self.parameters.withFrameSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_withFrameSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_withFrameLabel, self.kWidgetID_withFrameSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_withFrameSelector, "Select the column for the frame creation")
        vs.CreateCheckBox(self.dialog, self.kWidgetID_withFrame, "Include Frame    ")
        vs.SetBooleanItem(self.dialog, self.kWidgetID_withFrame, self.parameters.pictureParameters.withImage == "True")
        vs.SetRightItem(self.dialog, self.kWidgetID_withFrameSelector, self.kWidgetID_withFrame, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_withFrame, "Choose the value for the frame creation")
        # Frame Width dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameWidthLabel, "Frame Width: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_withFrameLabel, self.kWidgetID_frameWidthLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_frameWidthSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_frameWidthSelector,
                                                self.parameters.frameWidthSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameWidthSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameWidthLabel, self.kWidgetID_frameWidthSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameWidthSelector, "Select the column for the frame width")
        # Frame Height dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameHeightLabel, "Frame Height: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameWidthLabel, self.kWidgetID_frameHeightLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_frameHeightSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_frameHeightSelector,
                                                self.parameters.frameHeightSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameHeightSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameHeightLabel, self.kWidgetID_frameHeightSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameHeightSelector, "Select the column for the frame height")
        # Frame Thickness dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameThicknessLabel, "Frame Thickness: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameHeightLabel, self.kWidgetID_frameThicknessLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_frameThicknessSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_frameThicknessSelector,
                                                self.parameters.frameThicknessSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameThicknessSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameThicknessLabel, self.kWidgetID_frameThicknessSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameThicknessSelector, "Select the column for the frame thickness")
        valid, value = vs.ValidNumStr(self.parameters.pictureParameters.frameThickness)
        if not valid:
            value = PictureParameters().frameThickness
        vs.CreateEditReal(self.dialog, self.kWidgetID_frameThickness, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameThicknessSelector, self.kWidgetID_frameThickness, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameThickness, "Enter the thickness of the frame here.")
        # Frame Depth dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameDepthLabel, "Frame Depth: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameThicknessLabel, self.kWidgetID_frameDepthLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_frameDepthSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_frameDepthSelector,
                                                self.parameters.frameDepthSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameDepthSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameDepthLabel, self.kWidgetID_frameDepthSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameDepthSelector, "Select the column for the frame depth")
        valid, value = vs.ValidNumStr(self.parameters.pictureParameters.frameDepth)
        if not valid:
            value = PictureParameters().frameDepth
        vs.CreateEditReal(self.dialog, self.kWidgetID_frameDepth, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameDepthSelector, self.kWidgetID_frameDepth, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameDepth, "Enter the depth of the frame here.")
        # Frame Class
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameClassLabel, "Frame Class: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameDepthLabel, self.kWidgetID_frameClassLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_frameClassSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_frameClassSelector,
                                                self.parameters.frameClassSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameClassSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameClassLabel, self.kWidgetID_frameClassSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameClassSelector, "Select the column for the frame class")
        vs.CreateClassPullDownMenu(self.dialog, self.kWidgetID_frameClass, input_field_width)
        class_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_frameClass,
                                             self.parameters.pictureParameters.frameClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameClass, class_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameClassSelector, self.kWidgetID_frameClass, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameClass, "Enter the class of the frame here.")
        # Frame Texture scale
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameTextureScaleLabel, "Frame Texture Scale: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameClassLabel, self.kWidgetID_frameTextureScaleLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_frameTextureScaleSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_frameTextureScaleSelector,
                                                self.parameters.frameTextureScaleSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameTextureScaleSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameTextureScaleLabel,
                        self.kWidgetID_frameTextureScaleSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameTextureScaleSelector,
                       "Select the column for the frame texture scale")
        valid, value = vs.ValidNumStr(self.parameters.pictureParameters.frameTextureScale)
        if not valid:
            value = PictureParameters().frameTextureScale
        vs.CreateEditReal(self.dialog, self.kWidgetID_frameTextureScale, 1, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameTextureScaleSelector, self.kWidgetID_frameTextureScale, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameTextureScale, "Enter the frame texture scale")
        # Frame Texture rotation
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameTextureRotationLabel, "Frame Texture Rotation: ",
                            label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameTextureScaleLabel,
                        self.kWidgetID_frameTextureRotationLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_frameTextureRotationSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_frameTextureRotationSelector,
                                                self.parameters.frameTextureRotationSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameTextureRotationSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameTextureRotationLabel,
                        self.kWidgetID_frameTextureRotationSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameTextureRotationSelector,
                       "Select the column for the frame texture rotation")
        valid, value = vs.ValidNumStr(self.parameters.pictureParameters.frameTextureRotation)
        if not valid:
            value = PictureParameters().frameTextureRotation
        vs.CreateEditReal(self.dialog, self.kWidgetID_frameTextureRotation, 1, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameTextureRotationSelector,
                        self.kWidgetID_frameTextureRotation,
                        0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameTextureRotation, "Enter the frame texture scale")

        # Matboard group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_matboardGroup, "Matboard", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameGroup, self.kWidgetID_matboardGroup, 0, 0)

        # With Matboard checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_withMatboardLabel, "With Matboard: ", label_width)
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_matboardGroup, self.kWidgetID_withMatboardLabel)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_withMatboardSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_withMatboardSelector,
                                                self.parameters.withMatboardSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_withMatboardSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_withMatboardLabel, self.kWidgetID_withMatboardSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_withMatboardSelector, "Select the column for the Matboard creation")
        vs.CreateCheckBox(self.dialog, self.kWidgetID_withMatboard, "Include Matboard")
        vs.SetBooleanItem(self.dialog, self.kWidgetID_withMatboard,
                          self.parameters.pictureParameters.withMatboard == "True")
        vs.SetRightItem(self.dialog, self.kWidgetID_withMatboardSelector, self.kWidgetID_withMatboard, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_withMatboard, "Choose the value for the Matboard creation")

        # Matboard Position dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_matboardPositionLabel, "Matboard Position: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_withMatboardLabel, self.kWidgetID_matboardPositionLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_matboardPositionSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_matboardPositionSelector,
                                                self.parameters.matboardPositionSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardPositionSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardPositionLabel,
                        self.kWidgetID_matboardPositionSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardPositionSelector,
                       "Select the column for the matboard position")
        valid, value = vs.ValidNumStr(self.parameters.pictureParameters.matboardPosition)
        if not valid:
            value = PictureParameters().matboardPosition
        vs.CreateEditReal(self.dialog, self.kWidgetID_matboardPosition, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardPositionSelector, self.kWidgetID_matboardPosition, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardPosition, "Enter the position (depth) of the matboard here.")
        # Matboard Class
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_matboardClassLabel, "Matboard Class: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_matboardPositionLabel, self.kWidgetID_matboardClassLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_matboardClassSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_matboardClassSelector,
                                                self.parameters.matboardClassSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardClassSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardClassLabel, self.kWidgetID_matboardClassSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardClassSelector, "Select the column for the matboard class")
        vs.CreateClassPullDownMenu(self.dialog, self.kWidgetID_matboardClass, input_field_width)
        class_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_matboardClass,
                                             self.parameters.pictureParameters.matboardClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardClass, class_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardClassSelector, self.kWidgetID_matboardClass, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardClass, "Enter the class of the matboard here.")
        # Frame Texture scale
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_matboardTextureScaleLabel, "Matboard Texture Scale: ",
                            label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_matboardClassLabel, self.kWidgetID_matboardTextureScaleLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_matboardTextureScaleSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_matboardTextureScaleSelector,
                                                self.parameters.matboardTextureScaleSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardTextureScaleSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardTextureScaleLabel,
                        self.kWidgetID_matboardTextureScaleSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardTextureScaleSelector,
                       "Select the column for the matboard texture scale")
        valid, value = vs.ValidNumStr(self.parameters.pictureParameters.matboardTextureScale)
        if not valid:
            value = PictureParameters().matboardTextureScale
        vs.CreateEditReal(self.dialog, self.kWidgetID_matboardTextureScale, 1, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardTextureScaleSelector,
                        self.kWidgetID_matboardTextureScale, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardTextureScale, "Enter the matboard texture scale")
        # Frame Texture rotation
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_matboardTextureRotatLabel, "Matboard Texture Rotation: ",
                            label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_matboardTextureScaleLabel, self.kWidgetID_matboardTextureRotatLabel,
                        0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_matboardTextureRotatSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_matboardTextureRotatSelector,
                                                self.parameters.matboardTextureRotatSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardTextureRotatSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardTextureRotatLabel,
                        self.kWidgetID_matboardTextureRotatSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardTextureRotatSelector,
                       "Select the column for the matboard texture rotation")
        valid, value = vs.ValidNumStr(self.parameters.pictureParameters.matboardTextureRotat)
        if not valid:
            value = PictureParameters().matboardTextureRotat
        vs.CreateEditReal(self.dialog, self.kWidgetID_matboardTextureRotat, 1, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardTextureRotatSelector,
                        self.kWidgetID_matboardTextureRotat, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardTextureRotat, "Enter the matboard texture scale")

        # Glass group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_glassGroup, "Glass", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_matboardGroup, self.kWidgetID_glassGroup, 0, 0)

        # With Glass checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_withGlassLabel, "With Glass: ", label_width)
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_glassGroup, self.kWidgetID_withGlassLabel)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_withGlassSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_withGlassSelector,
                                                self.parameters.withGlassSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_withGlassSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_withGlassLabel, self.kWidgetID_withGlassSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_withGlassSelector, "Select the column for the Glass creation")
        vs.CreateCheckBox(self.dialog, self.kWidgetID_withGlass, "Include Galss    ")
        vs.SetBooleanItem(self.dialog, self.kWidgetID_withGlass, self.parameters.pictureParameters.withGlass == "True")
        vs.SetRightItem(self.dialog, self.kWidgetID_withGlassSelector, self.kWidgetID_withGlass, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_withGlass, "Choose the value for the Glass creation")
        # Glass Position dimension
        # -----------------------------------------------------------------------------------------

        vs.CreateStaticText(self.dialog, self.kWidgetID_glassPositionLabel, "Glass Position: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_withGlassLabel, self.kWidgetID_glassPositionLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_glassPositionSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_glassPositionSelector,
                                                self.parameters.glassPositionSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_glassPositionSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_glassPositionLabel, self.kWidgetID_glassPositionSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_glassPositionSelector, "Select the column for the glass position")
        valid, value = vs.ValidNumStr(self.parameters.pictureParameters.glassPosition)
        if not valid:
            value = PictureParameters().glassPosition
        vs.CreateEditReal(self.dialog, self.kWidgetID_glassPosition, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_glassPositionSelector, self.kWidgetID_glassPosition, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_glassPosition, "Enter the position (depth) of the glass here.")
        # Glass Class
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_glassClassLabel, "Glass Class: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_glassPositionLabel, self.kWidgetID_glassClassLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_glassClassSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_glassClassSelector,
                                                self.parameters.glassClassSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_glassClassSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_glassClassLabel, self.kWidgetID_glassClassSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_glassClassSelector, "Select the column for the glass class")
        vs.CreateClassPullDownMenu(self.dialog, self.kWidgetID_glassClass, input_field_width)
        class_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_glassClass,
                                             self.parameters.pictureParameters.glassClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_glassClass, class_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_glassClassSelector, self.kWidgetID_glassClass, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_glassClass, "Enter the class of the glass here.")

        # Criteria group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_excelCriteriaGroup, "Criteria", True)
        vs.SetRightItem(self.dialog, self.kWidgetID_imageGroup, self.kWidgetID_excelCriteriaGroup, 0, 0)
        # Criteria
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_excelCriteriaLabel, "Picture Creation Criteria: ", label_width)
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_excelCriteriaGroup, self.kWidgetID_excelCriteriaLabel)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_excelCriteriaSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_excelCriteriaSelector,
                                                self.parameters.excelCriteriaSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_excelCriteriaSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_excelCriteriaLabel, self.kWidgetID_excelCriteriaSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_excelCriteriaSelector, "Select the column for selection criteria")

        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_excelCriteriaValue, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_excelCriteriaValue,
                                                self.parameters.excelCriteriaValue)
        vs.SelectChoice(self.dialog, self.kWidgetID_excelCriteriaValue, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_excelCriteriaSelector, self.kWidgetID_excelCriteriaValue, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_excelCriteriaValue, "Select the selection criteria value")

        # Symbol group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_SymbolGroup, "Symbol", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_excelCriteriaGroup, self.kWidgetID_SymbolGroup, 0, 0)
        # Create Symbol checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_SymbolCreateSymbol, "Create Symbol")
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_SymbolGroup, self.kWidgetID_SymbolCreateSymbol)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_SymbolCreateSymbol, self.parameters.symbolCreateSymbol == "True")
        vs.SetHelpText(self.dialog, self.kWidgetID_SymbolCreateSymbol, "Check to create a symbol for every Picture")
        # Symbol Folder
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_SymbolFolderLabel, "Symbol Folder: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_SymbolCreateSymbol, self.kWidgetID_SymbolFolderLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_SymbolFolderSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_SymbolFolderSelector,
                                                self.parameters.symbolFolderSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_SymbolFolderSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_SymbolFolderLabel, self.kWidgetID_SymbolFolderSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_SymbolFolderSelector, "Select the column for the symbol folder name")

        vs.CreateEditText(self.dialog, self.kWidgetID_SymbolFolder, self.parameters.symbolFolder, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_SymbolFolderSelector, self.kWidgetID_SymbolFolder, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_SymbolFolder, "Enter the symbol folder name")

        # Import group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_importGroup, "Import", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_SymbolGroup, self.kWidgetID_importGroup, 0, 0)

        # Create missing classes
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_createMissingClasses, "Create missing classes")
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_importGroup, self.kWidgetID_createMissingClasses)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_createMissingClasses, self.parameters.createMissingClasses == "True")
        vs.SetHelpText(self.dialog, self.kWidgetID_createMissingClasses, "Create missing classes")

        # Ignore Existing
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_importIgnoreExisting, "Ignore manual fields on existing Pictures")
        vs.SetBelowItem(self.dialog, self.kWidgetID_createMissingClasses, self.kWidgetID_importIgnoreExisting, 0, 0)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_importIgnoreExisting, self.parameters.importIgnoreExisting == "True")
        vs.SetHelpText(self.dialog, self.kWidgetID_importIgnoreExisting, "Ignore manual fields on existing Pictures")
        # Ignore Errors
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_importIgnoreErrors, "Ignore Errors")
        vs.SetBelowItem(self.dialog, self.kWidgetID_importIgnoreExisting, self.kWidgetID_importIgnoreErrors, 0, 0)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_importIgnoreErrors, self.parameters.importIgnoreErrors == "True")
        vs.SetHelpText(self.dialog, self.kWidgetID_importIgnoreErrors, "Check to ignore all import errors")
        # Ignore Unmodified
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_importIgnoreUnmodified, "Ignore Unmodified")
        vs.SetBelowItem(self.dialog, self.kWidgetID_importIgnoreErrors, self.kWidgetID_importIgnoreUnmodified, 0, 0)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_importIgnoreUnmodified,
                          self.parameters.importIgnoreUnmodified == "True")
        vs.SetHelpText(self.dialog, self.kWidgetID_importIgnoreUnmodified, "Check to ignore all unmodified pictures")

        # Import Button
        # -----------------------------------------------------------------------------------------
        vs.CreatePushButton(self.dialog, self.kWidgetID_importButton, "Import")
        vs.SetBelowItem(self.dialog, self.kWidgetID_importIgnoreUnmodified, self.kWidgetID_importButton, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_fileBrowseButton, "Click to start the import operation")
        # New Pictures Count
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_importNewCount,
                            "New Pictures: {}".format(self.importNewCount), label_width + 10)
        vs.SetBelowItem(self.dialog, self.kWidgetID_importButton, self.kWidgetID_importNewCount, 0, 0)
        # Updated Pictures Count
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_importUpdatedCount,
                            "Updated Pictures: {}".format(self.importUpdatedCount), label_width + 10)
        vs.SetBelowItem(self.dialog, self.kWidgetID_importNewCount, self.kWidgetID_importUpdatedCount, 0, 0)
        # Deleted Pictures Count
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_importDeletedCount,
                            "Deleted Pictures: {}".format(self.importDeletedCount), label_width + 10)
        vs.SetBelowItem(self.dialog, self.kWidgetID_importUpdatedCount, self.kWidgetID_importDeletedCount, 0, 0)
        # Error Pictures Count
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_importErrorCount,
                            "Error Pictures: {}".format(self.importErrorCount), label_width + 10)
        vs.SetBelowItem(self.dialog, self.kWidgetID_importDeletedCount, self.kWidgetID_importErrorCount, 0, 0)

    def update_picture(self, picture_parameters: PictureParameters, log_file: IO):
        log_message = ""
        image_message = ""
        frame_message = ""
        matboard_message = ""
        glass_message = ""
        changed = False

        existing_picture = vs.GetObject(picture_parameters.pictureName)
        if self.parameters.withImageSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
            existing_with_image = vs.GetRField(existing_picture, "Picture", "WithImage")
            if picture_parameters.withImage != existing_with_image:
                if picture_parameters.withImage == "True":
                    image_message = "- Add immage " + image_message
                else:
                    image_message = "- Removed image "
                vs.SetRField(existing_picture, "Picture", "WithImage", picture_parameters.withImage)
                changed = True

        if picture_parameters.withImage == "True":
            valid, existing_image_width = vs.ValidNumStr(vs.GetRField(existing_picture, "Picture", "ImageWidth"))
            existing_image_width = round(existing_image_width, 3)
            if picture_parameters.imageWidth != existing_image_width:
                image_message = image_message + "- Image With changed "
                vs.SetRField(existing_picture, "Picture", "ImageWidth", picture_parameters.imageWidth)
                changed = True

            valid, existing_image_height = vs.ValidNumStr(vs.GetRField(existing_picture, "Picture", "ImageHeight"))
            existing_image_height = round(existing_image_height, 3)
            if picture_parameters.imageHeight != existing_image_height:
                image_message = image_message + "- Image Height changed "
                vs.SetRField(existing_picture, "Picture", "ImageHeight", picture_parameters.imageHeight)
                changed = True

            if self.parameters.imagePositionSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                valid, existing_image_position = vs.ValidNumStr(
                    vs.GetRField(existing_picture, "Picture", "ImagePosition"))
                existing_image_position = round(existing_image_position, 3)
                if picture_parameters.imagePosition != existing_image_position:
                    image_message = image_message + "- Image Position changed "
                    vs.SetRField(existing_picture, "Picture", "ImagePosition", picture_parameters.imagePosition)
                    changed = True

        if self.parameters.withFrameSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
            existing_with_frame = vs.GetRField(existing_picture, "Picture", "WithFrame")
            if picture_parameters.withFrame != existing_with_frame:
                if picture_parameters.withFrame == "True":
                    frame_message = "Add frame " + frame_message
                else:
                    frame_message = "Removed frame "
                vs.SetRField(existing_picture, "Picture", "WithFrame", picture_parameters.withFrame)
                changed = True

        if picture_parameters.withFrame == "True":
            valid, existing_frame_width = vs.ValidNumStr(vs.GetRField(existing_picture, "Picture", "FrameWidth"))
            existing_frame_width = round(existing_frame_width, 3)
            if picture_parameters.frameWidth != existing_frame_width:
                frame_message = frame_message + "- Frame Width changed "
                vs.SetRField(existing_picture, "Picture", "FrameWidth", picture_parameters.frameWidth)
                changed = True

            valid, existing_frame_height = vs.ValidNumStr(vs.GetRField(existing_picture, "Picture", "FrameHeight"))
            existing_frame_height = round(existing_frame_height, 3)
            if picture_parameters.frameHeight != existing_frame_height:
                frame_message = frame_message + "- Frame Height changed "
                vs.SetRField(existing_picture, "Picture", "FrameHeight", picture_parameters.frameHeight)
                changed = True

            if self.parameters.frameThicknessSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                valid, existing_frame_thickness = vs.ValidNumStr(
                    vs.GetRField(existing_picture, "Picture", "FrameThickness"))
                existing_frame_thickness = round(existing_frame_thickness, 3)
                if picture_parameters.frameThickness != existing_frame_thickness:
                    frame_message = frame_message + "- Frame Thickness changed "
                    vs.SetRField(existing_picture, "Picture", "FrameThickness", picture_parameters.frameThickness)
                    changed = True

            if self.parameters.frameDepthSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                valid, existing_frame_depth = vs.ValidNumStr(vs.GetRField(existing_picture, "Picture", "FrameDepth"))
                existing_frame_depth = round(existing_frame_depth, 3)
                if picture_parameters.frameDepth != existing_frame_depth:
                    frame_message = frame_message + "- Frame Depth changed "
                    vs.SetRField(existing_picture, "Picture", "FrameDepth", picture_parameters.frameDepth)
                    changed = True

            if self.parameters.frameClassSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                existing_frame_class = vs.GetRField(existing_picture, "Picture", "FrameClass")
                if picture_parameters.frameClass != existing_frame_class:
                    frame_message = frame_message + "- Frame Class changed "
                    vs.SetRField(existing_picture, "Picture", "FrameClass", picture_parameters.frameClass)
                    changed = True

            if self.parameters.frameTextureScaleSelector != "-- Manual" \
                    or self.parameters.importIgnoreExisting == "False":
                valid, existing_frame_texture_scale = vs.ValidNumStr(
                    vs.GetRField(existing_picture, "Picture", "FrameTextureScale"))
                existing_frame_texture_scale = round(existing_frame_texture_scale, 3)
                if picture_parameters.frameTextureScale != existing_frame_texture_scale:
                    frame_message = frame_message + "- Frame Texture Scale changed "
                    vs.SetRField(existing_picture, "Picture", "FrameTextureScale", picture_parameters.frameTextureScale)
                    changed = True

            if self.parameters.frameTextureRotationSelector != "-- Manual" \
                    or self.parameters.importIgnoreExisting == "False":
                valid, existing_frame_texture_rotation = vs.ValidNumStr(
                    vs.GetRField(existing_picture, "Picture", "FrameTextureRotation"))
                existing_frame_texture_rotation = round(existing_frame_texture_rotation, 3)
                if picture_parameters.frameTextureRotation != existing_frame_texture_rotation:
                    frame_message = frame_message + "- Frame Texture Rotation changed "
                    vs.SetRField(existing_picture, "Picture", "FrameTextureRotation",
                                 picture_parameters.frameTextureRotation)
                    changed = True

        if self.parameters.withMatboardSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
            existing_with_matboard = vs.GetRField(existing_picture, "Picture", "WithMatboard")
            if picture_parameters.withMatboard != existing_with_matboard:
                if picture_parameters.withMatboard == "True":
                    matboard_message = "Add matboard " + matboard_message
                else:
                    matboard_message = "Removed matboard "
                vs.SetRField(existing_picture, "Picture", "WithMatboard", picture_parameters.withMatboard)
                changed = True

        if picture_parameters.withMatboard == "True":
            valid, existing_frame_width = vs.ValidNumStr(vs.GetRField(existing_picture, "Picture", "FrameWidth"))
            existing_frame_width = round(existing_frame_width, 3)
            if picture_parameters.frameWidth != existing_frame_width:
                frame_message = frame_message + "- Frame Width changed "
                vs.SetRField(existing_picture, "Picture", "FrameWidth", picture_parameters.frameWidth)
                changed = True

            valid, existing_frame_height = vs.ValidNumStr(vs.GetRField(existing_picture, "Picture", "FrameHeight"))
            existing_frame_height = round(existing_frame_height, 3)
            if picture_parameters.frameHeight != existing_frame_height:
                frame_message = frame_message + "- Frame Height changed "
                vs.SetRField(existing_picture, "Picture", "FrameHeight", picture_parameters.frameHeight)
                changed = True

            if self.parameters.matboardPositionSelector != "-- Manual" \
                    or self.parameters.importIgnoreExisting == "False":
                valid, existing_matboard_position = vs.ValidNumStr(
                    vs.GetRField(existing_picture, "Picture", "MatboardPosition"))
                existing_matboard_position = round(existing_matboard_position, 3)
                if picture_parameters.matboardPosition != existing_matboard_position:
                    matboard_message = matboard_message + "- Matboard Position changed "
                    vs.SetRField(existing_picture, "Picture", "MatboardPosition", picture_parameters.matboardPosition)
                    changed = True

            if self.parameters.matboardClassSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                existing_matboard_class = vs.GetRField(existing_picture, "Picture", "MatboardClass")
                if picture_parameters.matboardClass != existing_matboard_class:
                    matboard_message = matboard_message + "- Matboard Class changed "
                    vs.SetRField(existing_picture, "Picture", "MatboardClass", picture_parameters.matboardClass)
                    changed = True

            if self.parameters.matboardTextureScaleSelector != "-- Manual" \
                    or self.parameters.importIgnoreExisting == "False":
                valid, existing_matboard_texture_scale = vs.ValidNumStr(
                    vs.GetRField(existing_picture, "Picture", "MatboardTextureScale"))
                existing_matboard_texture_scale = round(existing_matboard_texture_scale, 3)
                if picture_parameters.matboardTextureScale != existing_matboard_texture_scale:
                    matboard_message = matboard_message + "- Matboard Texture Scale changed "
                    vs.SetRField(existing_picture, "Picture", "MatboardTextureScale",
                                 picture_parameters.matboardTextureScale)
                    changed = True

            if self.parameters.matboardTextureRotatSelector != "-- Manual" \
                    or self.parameters.importIgnoreExisting == "False":
                valid, existing_matboard_texture_rotat = vs.ValidNumStr(
                    vs.GetRField(existing_picture, "Picture", "MatboardTextureRotat"))
                existing_matboard_texture_rotat = round(existing_matboard_texture_rotat, 3)
                if picture_parameters.matboardTextureRotat != existing_matboard_texture_rotat:
                    matboard_message = matboard_message + "- Matboard Texture Rotation changed "
                    vs.SetRField(existing_picture, "Picture", "MatboardTextureRotat",
                                 picture_parameters.matboardTextureRotat)
                    changed = True

        if self.parameters.withGlassSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
            existing_with_glass = vs.GetRField(existing_picture, "Picture", "WithGlass")
            if picture_parameters.withGlass != existing_with_glass:
                if picture_parameters.withGlass == "True":
                    glass_message = "Add glass " + image_message
                else:
                    glass_message = "Removed glass "
                vs.SetRField(existing_picture, "Picture", "WithGlass", picture_parameters.withGlass)
                changed = True

        if picture_parameters.withGlass == "True":
            if self.parameters.glassPositionSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                valid, existing_glass_position = vs.ValidNumStr(
                    vs.GetRField(existing_picture, "Picture", "GlassPosition"))
                existing_glass_position = round(existing_glass_position, 3)
                if picture_parameters.glassPosition != existing_glass_position:
                    glass_message = glass_message + "- Glass Position changed "
                    vs.SetRField(existing_picture, "Picture", "GlassPosition", picture_parameters.glassPosition)
                    changed = True

            if self.parameters.glassClassSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                existing_glass_class = vs.GetRField(existing_picture, "Picture", "GlassClass")
                if picture_parameters.glassClass != existing_glass_class:
                    glass_message = glass_message + "- Glass Class changed "
                    vs.SetRField(existing_picture, "Picture", "GlassClass", picture_parameters.glassClass)
                    changed = True

        if changed:
            vs.ResetObject(existing_picture)

            log_message = "{} * [Modified] ".format(
                picture_parameters.pictureName) +\
                image_message + frame_message + matboard_message + glass_message + "\n"
            self.importUpdatedCount += 1
        else:
            if self.parameters.importIgnoreUnmodified != "True":
                log_message = "{} * [Unmodified] \n".format(picture_parameters.pictureName)

        if log_message:
            log_file.write(log_message)

    def new_picture(self, picture_parameters: PictureParameters, log_file: IO):
        if picture_parameters.withImage == "True" \
                or picture_parameters.withFrame == "True" \
                or picture_parameters.withMatboard == "True" \
                or picture_parameters.withGlass == "True":

            texture = vs.GetObject("Arroway {}".format(picture_parameters.pictureName.replace('-', ' ').replace('_', ' ')))
            if texture == 0:
                for outer in range(0, 99):
                    for inner in range(1, 99):
                        if outer == 0:
                            texture_name = "Arroway {}".format(picture_parameters.pictureName.replace('-', ' ').replace('_', ' ')) + ' ' + str(inner)
                        else:
                            texture_name = "Arroway {}".format(picture_parameters.pictureName.replace('-', ' ').replace('_', ' ')) + ' ' + str(inner) + ' ' + str(outer)
                        texture = vs.GetObject(texture_name)
                        if texture != 0:
                            break
                    if texture != 0:
                        break
            if texture == 0:
                picture_parameters.imageTexture = ""
                log_message = "{} * [Could not find texture] \n".format(picture_parameters.pictureName)
                self.importErrorCount += 1
            else:
                # Create a new Picture Object
                active_class = vs.ActiveClass()
                vs.NameClass("Pictures")
                picture_parameters.imageTexture = vs.GetName(texture)
                build_picture(picture_parameters, None)

                log_message = "{} * [New] \n".format(picture_parameters.pictureName)
                self.importNewCount += 1
                vs.NameClass(active_class)

            log_file.write(log_message)

    def import_pictures(self):
        vs.ProgressDlgOpen("Importing Pictures", True)
        total_rows = self.excel.get_worksheet_row_count()
        vs.ProgressDlgSetMeter("Importing " + str(total_rows) + " Pictures ...")
        vs.ProgressDlgStart(100.0, total_rows)
        self.importNewCount = 0
        self.importUpdatedCount = 0
        self.importDeletedCount = 0
        self.importErrorCount = 0
        document_file_name = vs.GetFPathName()
        document_folder = os.path.dirname(document_file_name)
        if not document_folder:
            document_folder = "C:/tmp"
        log_file_name = document_folder + "/" + "Import_Pictures_" + strftime("%y_%m_%d_%H_%M_%S", gmtime()) + ".log"

        log_file = open(log_file_name, "w")

        for picture_parameters in self.excel.get_worksheet_rows(log_file):

            if vs.ProgressDlgHasCancel():
                break
            vs.ProgressDlgYield(1)
            vs.ProgressDlgSetTopMsg("New Pictures: {}".format(self.importNewCount))
            vs.ProgressDlgSetBotMsg("Modified Pictures: {}".format(self.importUpdatedCount))

            if picture_parameters.pictureName:
                # self.set_texture(picture_parameters)
                existing_picture = vs.GetObject(picture_parameters.pictureName)
                if existing_picture:
                    self.update_picture(picture_parameters, log_file)
                else:
                    self.new_picture(picture_parameters, log_file)
            else:
                pass

        vs.ProgressDlgEnd()
        vs.ProgressDlgClose()

        log_file.write("--------------------------------------------------------------------------\n")
        log_file.write("Total new Pictures: {}\n".format(self.importNewCount))
        log_file.write("Total modified Pictures: {}\n".format(self.importUpdatedCount))
        log_file.write("Total deleted Pictures: {}\n".format(self.importDeletedCount))
        if self.parameters.importIgnoreErrors != "True":
            log_file.write("Total error Pictures: {}\n".format(self.importErrorCount))
        log_file.write("--------------------------------------------------------------------------\n")
        log_file.close()

    # def set_texture(self, picture_parameters: PictureParameters):
    #
    #     texture_name = "{} Picture Texture".format(picture_parameters.pictureName)
    #     texture = vs.GetObject(texture_name)
    #     if not texture:
    #         image_folder = os.path.join(os.path.join(os.path.dirname(self.parameters.excelFileName), "Images"), picture_parameters.symbolFolder)
    #         image_file = os.path.join(image_folder, "{}.jpg".format(picture_parameters.pictureName))
    #         image = vs.ImportImageFile(image_file, (10, 10))
    #         if not image:
    #             image = vs.FSActLayer()
    #         if image:
    #             paint = vs.CreatePaintFromImgN(image, (0, 0), 0)
    #             paint = vs.FSActLayer()
    #             if paint:
    #                 bitmap = vs.CreateTextureBitmap()
    #                 if bitmap:
    #                     vs.SetTexBitPaintNode(bitmap, paint)
    #                     vs.SetName(texture, texture_name)
    #                 else:
    #                     vs.DelObject(paint)
    #             else:
    #                 vs.DelObject(image)
