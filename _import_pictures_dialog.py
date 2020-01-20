"""
Created on Dec 7, 2016

@author: Carles
"""
import os
from time import strftime, gmtime
from typing import IO
import vs
from vs_constants import *
from _import_settings import ImportSettings
from _import_picture_database import ImportDatabase
from _picture_settings import PictureParameters, PictureRecord
from _picture import build_picture

# import pydevd_pycharm
# pydevd_pycharm.settrace('localhost', port=12345, stdoutToServer=True, stderrToServer=True, suspend=False)


def same_dimension(dimemsion_1: str, dimension_2: str) -> bool:
    dim1 = 0
    dim2 = 0
    valid1, value = vs.ValidNumStr(dimemsion_1)
    if valid1:
        dim1 = round(value, 3)
    valid2, value = vs.ValidNumStr(dimension_2)
    if valid2:
        dim2 = round(value, 3)
    return valid1 == valid2 and dim1 == dim2


def dimension_strings(dimension_1: str, dimension_2: str) -> (str, str):
    valid1, value = vs.ValidNumStr(dimension_1)
    if valid1:
        dim1 = round(value, 3)
        str1 = "{}".format(dim1)
    else:
        str1 = "Invalid"

    valid2, value = vs.ValidNumStr(dimension_2)
    if valid2:
        dim2 = round(value, 3)
        str2 = "{}".format(dim2)
    else:
        str2 = "Invalid"
    return str1, str2


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
        self.kWidgetID_windowWidthLabel = 86
        self.kWidgetID_windowWidthSelector = 87
        self.kWidgetID_windowHeightLabel = 88
        self.kWidgetID_windowHeightSelector = 89
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
        # Full, no more id's

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
        self.kWidgetID_symbolGroup = 200
        self.kWidgetID_symbolCreateSymbol = 201
        self.kWidgetID_symbolFolderLabel = 202
        self.kWidgetID_symbolFolderSelector = 203
        self.kWidgetID_symbolFolder = 204

        # Use Classes
        self.kWidgetID_classGroup = 300
        self.kWidgetID_classAssignPictureClass = 301
        self.kWidgetID_classPictureClassLabel = 302
        self.kWidgetID_classPictureClassSelector = 303
        self.kWidgetID_classPictureClass = 304
        self.kWidgetID_classCreateMissingClasses = 305

        # Import Operation
        self.kWidgetID_importGroup = 400
        self.kWidgetID_importIgnoreErrors = 401
        self.kWidgetID_importIgnoreExisting = 402
        self.kWidgetID_importIgnoreUnmodified = 403
        self.kWidgetID_importButton = 404
        self.kWidgetID_importNewCount = 405
        self.kWidgetID_importUpdatedCount = 406
        self.kWidgetID_importDeletedCount = 407
        self.kWidgetID_importErrorCount = 408

        # Metadata
        self.kWidgetID_metaGroup = 500
        self.kWidgetID_metaImportMetadata = 550
        self.kWidgetID_metaArtworkTitleLabel = 501
        self.kWidgetID_metaArtworkTitleSelector = 502
        self.kWidgetID_metaAuthorNameLabel = 503
        self.kWidgetID_metaAuthorNameSelector = 504
        self.kWidgetID_metaArtworkCreationDateLabel = 505
        self.kWidgetID_metaArtworkCreationDateSelector = 506
        self.kWidgetID_metaArtworkMediaLabel = 507
        self.kWidgetID_metaArtworkMediaSelector = 508
        # self.kWidgetID_metaTypeLabel = 509
        # self.kWidgetID_metaTypeSelector = 510
        self.kWidgetID_metaRoomLocationLabel = 511
        self.kWidgetID_metaRoomLocationSelector = 512
        self.kWidgetID_metaArtworkSourceLabel = 513
        self.kWidgetID_metaArtworkSourceSelector = 514
        self.kWidgetID_metaRegistrationNumberLabel = 515
        self.kWidgetID_metaRegistrationNumberSelector = 516
        self.kWidgetID_metaAuthorBirthCountryLabel = 517
        self.kWidgetID_metaAuthorBirthCountrySelector = 518
        self.kWidgetID_metaAuthorBirthDateLabel = 519
        self.kWidgetID_metaAuthorBirthDateSelector = 520
        self.kWidgetID_metaAuthorDeathDateLabel = 521
        self.kWidgetID_metaAuthorDeathDateSelector = 522
        self.kWidgetID_metaDesignNotesLabel = 523
        self.kWidgetID_metaDesignNotesSelector = 524
        self.kWidgetID_metaExhibitionMediaLabel = 525
        self.kWidgetID_metaExhibitionMediaSelector = 526

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
            self.empty_all_fields()
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

        if state is True and self.parameters.excelCriteriaSelector != "-- Select column ...":
            criteria_values = self.excel.get_criteria_values()
            if criteria_values:
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
            self.parameters.imageTextureSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_withImageSelector, data)
        elif item == self.kWidgetID_imageWidthSelector:
            self.parameters.imageWidthSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_imageWidthSelector, data)
        elif item == self.kWidgetID_imageHeightSelector:
            self.parameters.imageHeightSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_imageHeightSelector, data)
        elif item == self.kWidgetID_imagePositionSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_imagePosition, data == 0)
            self.parameters.imagePositionSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_imagePositionSelector, data)
        elif item == self.kWidgetID_imagePosition:
            self.parameters.pictureParameters.imagePosition = str(vs.GetEditReal(self.dialog, self.kWidgetID_imagePosition, 3))
        elif item == self.kWidgetID_withFrameSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_withFrame, data == 0)
            self.parameters.withFrameSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_withFrameSelector, data)
        elif item == self.kWidgetID_withFrame:
            self.parameters.pictureParameters.withFrame = "{}".format(data != 0)
        elif item == self.kWidgetID_frameWidthSelector:
            self.parameters.frameWidthSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_frameWidthSelector, data)
        elif item == self.kWidgetID_frameHeightSelector:
            self.parameters.frameHeightSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_frameHeightSelector, data)
        elif item == self.kWidgetID_frameThicknessSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_frameThickness, data == 0)
            self.parameters.frameThicknessSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_frameThicknessSelector, data)
        elif item == self.kWidgetID_frameThickness:
            self.parameters.pictureParameters.frameThickness = str(vs.GetEditReal(
                self.dialog, self.kWidgetID_frameThickness, 3))
        elif item == self.kWidgetID_frameDepthSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_frameDepth, data == 0)
            self.parameters.frameDepthSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_frameDepthSelector, data)
        elif item == self.kWidgetID_frameDepth:
            self.parameters.pictureParameters.frameDepth = str(vs.GetEditReal(self.dialog, self.kWidgetID_frameDepth, 3))
        elif item == self.kWidgetID_frameClassSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_frameClass, data == 0)
            self.parameters.frameClassSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_frameClassSelector, data)
        elif item == self.kWidgetID_frameClass:
            index, self.parameters.pictureParameters.frameClass = vs.GetSelectedChoiceInfo(self.dialog, self.kWidgetID_frameClass, 0)
        elif item == self.kWidgetID_frameTextureScaleSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScale, data == 0)
            self.parameters.frameTextureScaleSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_frameTextureScaleSelector, data)
        elif item == self.kWidgetID_frameTextureScale:
            self.parameters.pictureParameters.frameTextureScale = str(vs.GetEditReal(self.dialog, self.kWidgetID_frameTextureScale, 1))
        elif item == self.kWidgetID_frameTextureRotationSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotation, data == 0)
            self.parameters.frameTextureRotationSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_frameTextureRotationSelector, data)
        elif item == self.kWidgetID_frameTextureRotation:
            self.parameters.pictureParameters.frameTextureRotation = str(vs.GetEditReal(self.dialog, self.kWidgetID_frameTextureRotation, 1))
        elif item == self.kWidgetID_withMatboardSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_withMatboard, data == 0)
            self.parameters.withMatboardSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_withMatboardSelector, data)
        elif item == self.kWidgetID_withMatboard:
            self.parameters.pictureParameters.withMatboard = "{}".format(data != 0)
        elif item == self.kWidgetID_matboardPositionSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_matboardPosition, data == 0)
            self.parameters.matboardPositionSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_matboardPositionSelector, data)
        elif item == self.kWidgetID_windowWidthSelector:
            self.parameters.windowWidthSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_windowWidthSelector, data)
        elif item == self.kWidgetID_windowHeightSelector:
            self.parameters.windowHeightSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_windowHeightSelector, data)
        elif item == self.kWidgetID_matboardPosition:
            self.parameters.pictureParameters.matboardPosition = str(vs.GetEditReal(self.dialog, self.kWidgetID_matboardPosition, 3))
        elif item == self.kWidgetID_matboardClassSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_matboardClass, data == 0)
            self.parameters.matboardClassSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_matboardClassSelector, data)
        elif item == self.kWidgetID_matboardClass:
            index, self.parameters.pictureParameters.matboardClass = vs.GetSelectedChoiceInfo(self.dialog, self.kWidgetID_matboardClass, 0)
        elif item == self.kWidgetID_matboardTextureScaleSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScale, data == 0)
            self.parameters.matboardTextureScaleSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_matboardTextureScaleSelector, data)
        elif item == self.kWidgetID_matboardTextureScale:
            self.parameters.pictureParameters.matboardTextureScale = str(vs.GetEditReal(self.dialog, self.kWidgetID_matboardTextureScale, 1))
        elif item == self.kWidgetID_matboardTextureRotatSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotat, data == 0)
            self.parameters.matboardTextureRotatSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_matboardTextureRotatSelector, data)
        elif item == self.kWidgetID_matboardTextureRotat:
            self.parameters.pictureParameters.matboardTextureRotat = str(vs.GetEditReal(self.dialog, self.kWidgetID_matboardTextureRotat, 1))
        elif item == self.kWidgetID_withGlassSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_withGlass, data == 0)
            self.parameters.withGlassSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_withGlassSelector, data)
        elif item == self.kWidgetID_withGlass:
            self.parameters.pictureParameters.withGlass = "{}".format(data != 0)
        elif item == self.kWidgetID_glassPositionSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_glassPosition, data == 0)
            self.parameters.glassPositionSelector = vs.GetChoiceText(
                self.dialog, self.kWidgetID_glassPositionSelector, data)
        elif item == self.kWidgetID_glassPosition:
            self.parameters.pictureParameters.glassPosition = str(vs.GetEditReal(self.dialog, self.kWidgetID_glassPosition, 3))
        elif item == self.kWidgetID_glassClassSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_glassClass, data == 0)
            self.parameters.glassClassSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_glassClassSelector, data)
        elif item == self.kWidgetID_glassClass:
            index, self.parameters.pictureParameters.glassClass = vs.GetSelectedChoiceInfo(self.dialog, self.kWidgetID_glassClass, 0)
        elif item == self.kWidgetID_excelCriteriaSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_excelCriteriaValue, data != 0)
            new_excel_criteria_selector = vs.GetChoiceText(self.dialog, self.kWidgetID_excelCriteriaSelector, data)
            if new_excel_criteria_selector != self.parameters.excelCriteriaSelector:
                self.parameters.excelCriteriaSelector = new_excel_criteria_selector
                self.update_criteria_values(False)
                if data != 0:
                    self.update_criteria_values(True)
                else:
                    index = vs.GetChoiceIndex(self.dialog, self.kWidgetID_excelCriteriaValue, self.parameters.excelCriteriaValue)
                    if index == -1:
                        vs.SelectChoice(self.dialog, self.kWidgetID_excelCriteriaValue, 0, True)
                        self.parameters.excelCriteriaValue = "Select a value ..."
                    else:
                        vs.SelectChoice(self.dialog, self.kWidgetID_excelCriteriaValue, index, True)
        elif item == self.kWidgetID_excelCriteriaValue:
            self.parameters.excelCriteriaValue = vs.GetChoiceText(self.dialog, self.kWidgetID_excelCriteriaValue, data)
        elif item == self.kWidgetID_symbolCreateSymbol:
            self.parameters.symbolCreateSymbol = "{}".format(data != 0)
            selector_index = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_symbolFolderSelector, 0)
            vs.EnableItem(self.dialog, self.kWidgetID_symbolFolderSelector, data)
            vs.EnableItem(self.dialog, self.kWidgetID_symbolFolder, selector_index == 0 and data == 1)
        elif item == self.kWidgetID_symbolFolderSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_symbolFolder, data == 0)
            self.parameters.symbolFolderSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_symbolFolderSelector, data)
        elif item == self.kWidgetID_classAssignPictureClass:
            self.parameters.classAssignPictureClass = "{}".format(data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_classPictureClassSelector, data == 1)
            selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_classPictureClassSelector, self.parameters.classClassPictureSelector)
            vs.EnableItem(self.dialog, self.kWidgetID_classPictureClass, selector_index == 0 and data != 0)
        elif item == self.kWidgetID_classPictureClassSelector:
            vs.EnableItem(self.dialog, self.kWidgetID_classPictureClass, data == 0)
            self.parameters.classClassPictureSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_classPictureClassSelector, data)
        elif item == self.kWidgetID_classPictureClass:
            index, self.parameters.pictureParameters.pictureClass = vs.GetSelectedChoiceInfo(self.dialog, self.kWidgetID_classPictureClass, 0)
        elif item == self.kWidgetID_classCreateMissingClasses:
            self.parameters.createMissingClasses = "{}".format(data == 1)
        elif item == self.kWidgetID_metaImportMetadata:
            self.parameters.metaImportMetadata = "{}".format(data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkTitleSelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorNameSelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkMediaSelector, data == 1)
            # vs.EnableItem(self.dialog, self.kWidgetID_metaTypeSelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaRoomLocationSelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkSourceSelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaRegistrationNumberSelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaDesignNotesSelector, data == 1)
            vs.EnableItem(self.dialog, self.kWidgetID_metaExhibitionMediaSelector, data == 1)
        elif item == self.kWidgetID_metaArtworkTitleSelector:
            self.parameters.metaArtworkTitleSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaArtworkTitleSelector, data)
        elif item == self.kWidgetID_metaAuthorNameSelector:
            self.parameters.metaAuthorNameSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaAuthorNameSelector, data)
        elif item == self.kWidgetID_metaArtworkCreationDateSelector:
            self.parameters.metaArtworkCreationDateSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector, data)
        elif item == self.kWidgetID_metaArtworkMediaSelector:
            self.parameters.metaArtworkMediaSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaArtworkMediaSelector, data)
        # elif item == self.kWidgetID_metaTypeSelector:
        #     self.parameters.metaTypeSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaTypeSelector, data)
        elif item == self.kWidgetID_metaRoomLocationSelector:
            self.parameters.metaRoomLocationSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaRoomLocationSelector, data)
        elif item == self.kWidgetID_metaArtworkSourceSelector:
            self.parameters.metaArtworkSourceSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaArtworkSourceSelector, data)
        elif item == self.kWidgetID_metaRegistrationNumberSelector:
            self.parameters.metaRegistrationNumberSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaRegistrationNumberSelector, data)
        elif item == self.kWidgetID_metaAuthorBirthCountrySelector:
            self.parameters.metaAuthorBirthCountrySelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector, data)
        elif item == self.kWidgetID_metaAuthorBirthDateSelector:
            self.parameters.metaAuthorBirthDateSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector, data)
        elif item == self.kWidgetID_metaAuthorDeathDateSelector:
            self.parameters.metaAuthorDeathDateSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector, data)
        elif item == self.kWidgetID_metaDesignNotesSelector:
            self.parameters.metaDesignNotesSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaDesignNotesSelector, data)
        elif item == self.kWidgetID_metaExhibitionMediaSelector:
            self.parameters.metaExhibitionMediaSelector = vs.GetChoiceText(self.dialog, self.kWidgetID_metaExhibitionMediaSelector, data)
        elif item == self.kWidgetID_importIgnoreErrors:
            self.parameters.importIgnoreErrors = "{}".format(data != 0)
            vs.ShowItem(self.dialog, self.kWidgetID_importErrorCount, data == 0)
        elif item == self.kWidgetID_importIgnoreExisting:
            self.parameters.importIgnoreExisting = "{}".format(data != 0)
        elif item == self.kWidgetID_importIgnoreUnmodified:
            self.parameters.importIgnoreUnmodified = "{}".format(data != 0)
        elif item == self.kWidgetID_importButton:
            self.import_pictures()
            vs.SetItemText(self.dialog, self.kWidgetID_importNewCount, "New Pictures: {}".format(self.importNewCount))
            vs.SetItemText(self.dialog, self.kWidgetID_importUpdatedCount, "Updated Pictures: {}".format(self.importUpdatedCount))
            vs.SetItemText(self.dialog, self.kWidgetID_importDeletedCount, "Deleted Pictures: {}".format(self.importDeletedCount))
            vs.SetItemText(self.dialog, self.kWidgetID_importErrorCount, "Error Pictures: {}".format(self.importErrorCount))

        # This section handles the following cases:
        # - The Dialog is initializing
        # - The name of the workbook file has changed
        if item == self.kWidgetID_fileName or item == self.kWidgetID_fileBrowseButton or item == KDialogInitEvent:
            self.set_workbook()

        # The image selection has changed
        if item == self.kWidgetID_withImageSelector or item == self.kWidgetID_withImage or item == self.kWidgetID_excelSheetName:
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
        if item == self.kWidgetID_withFrameSelector or item == self.kWidgetID_withFrame or item == self.kWidgetID_excelSheetName:
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
        if item == self.kWidgetID_withMatboardSelector or item == self.kWidgetID_withMatboard or item == self.kWidgetID_excelSheetName:
            state = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_withMatboardSelector, 0) != 0 or \
                    vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard) is True

            vs.EnableItem(self.dialog, self.kWidgetID_windowWidthLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_windowWidthSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_windowHeightLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_windowHeightSelector, state)
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
        if item == self.kWidgetID_withGlassSelector or item == self.kWidgetID_withGlass or item == self.kWidgetID_excelSheetName:
            state = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_withGlassSelector, 0) != 0 or \
                    vs.GetBooleanItem(self.dialog, self.kWidgetID_withGlass) is True

            vs.EnableItem(self.dialog, self.kWidgetID_glassPositionLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_glassPositionSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_glassPosition, state)
            vs.EnableItem(self.dialog, self.kWidgetID_glassClassLabel, state)
            vs.EnableItem(self.dialog, self.kWidgetID_glassClassSelector, state)
            vs.EnableItem(self.dialog, self.kWidgetID_glassClass, state)

        # After the event has been handled, update some of the import validity settings accordingly
        self.parameters.imageValid = ((self.parameters.withImageSelector == "-- Manual" and self.parameters.pictureParameters.withImage == "True") or
                                      self.parameters.withImageSelector != "-- Manual") and \
                                     (self.parameters.imageTextureSelector != "-- Select column ...") and \
                                     (self.parameters.imageWidthSelector != "-- Select column ...") and \
                                     (self.parameters.imageHeightSelector != "-- Select column ...")

        self.parameters.frameValid = ((self.parameters.withFrameSelector == "-- Manual" and self.parameters.pictureParameters.withFrame == "True") or
                                      self.parameters.withFrameSelector != "-- Manual") and \
                                     (self.parameters.frameWidthSelector != "-- Select column ...") and \
                                     (self.parameters.frameHeightSelector != "-- Select column ...")

        self.parameters.matboardValid = ((self.parameters.withMatboardSelector == "-- Manual" and self.parameters.pictureParameters.withMatboard == "True") or
                                         self.parameters.withMatboardSelector != "-- Manual") and \
                                        (self.parameters.windowWidthSelector != "-- Select column ...") and \
                                        (self.parameters.windowHeightSelector != "-- Select column ...")

        self.parameters.glassValid = ((self.parameters.withGlassSelector == "-- Manual" and
                                       self.parameters.pictureParameters.withGlass == "True") or self.parameters.withGlassSelector != "-- Manual")

        self.parameters.criteriaValid = \
            (self.parameters.excelCriteriaSelector != "-- Select column ..." and self.parameters.excelCriteriaValue != "Select a value ...")

        self.parameters.importValid = (self.parameters.imageValid or self.parameters.frameValid) and self.parameters.criteriaValid

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
        vs.ShowItem(self.dialog, self.kWidgetID_windowWidthLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_windowWidthSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_windowHeightLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_windowHeightSelector, state)
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

        vs.ShowItem(self.dialog, self.kWidgetID_symbolGroup, state)
        vs.ShowItem(self.dialog, self.kWidgetID_symbolCreateSymbol, state)
        vs.ShowItem(self.dialog, self.kWidgetID_symbolFolderLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_symbolFolderSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_symbolFolder, state)

        vs.ShowItem(self.dialog, self.kWidgetID_classGroup, state)
        vs.ShowItem(self.dialog, self.kWidgetID_classAssignPictureClass, state)
        vs.ShowItem(self.dialog, self.kWidgetID_classPictureClassLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_classPictureClassSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_classPictureClass, state)
        vs.ShowItem(self.dialog, self.kWidgetID_classCreateMissingClasses, state)

        vs.ShowItem(self.dialog, self.kWidgetID_metaArtworkTitleLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaArtworkTitleSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaAuthorNameLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaAuthorNameSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaArtworkCreationDateLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaArtworkMediaLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaArtworkMediaSelector, state)
        # vs.ShowItem(self.dialog, self.kWidgetID_metaTypeLabel, state)
        # vs.ShowItem(self.dialog, self.kWidgetID_metaTypeSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaRoomLocationLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaRoomLocationSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaArtworkSourceLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaArtworkSourceSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaRegistrationNumberLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaRegistrationNumberSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaAuthorBirthCountryLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaAuthorBirthDateLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaAuthorDeathDateLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaDesignNotesLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaDesignNotesSelector, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaExhibitionMediaLabel, state)
        vs.ShowItem(self.dialog, self.kWidgetID_metaExhibitionMediaSelector, state)

        vs.ShowItem(self.dialog, self.kWidgetID_importGroup, state)
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
            vs.AddChoice(self.dialog, self.kWidgetID_windowWidthSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_windowHeightSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_matboardPositionSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_matboardClassSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_matboardTextureScaleSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_matboardTextureRotatSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_withGlassSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_glassPositionSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_glassClassSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_excelCriteriaSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_symbolFolderSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_classPictureClassSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaArtworkTitleSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaAuthorNameSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaArtworkMediaSelector, column, 0)
            # vs.AddChoice(self.dialog, self.kWidgetID_metaTypeSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaRoomLocationSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaArtworkSourceSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaRegistrationNumberSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaDesignNotesSelector, column, 0)
            vs.AddChoice(self.dialog, self.kWidgetID_metaExhibitionMediaSelector, column, 0)

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
        vs.AddChoice(self.dialog, self.kWidgetID_windowWidthSelector, "-- Select column ...", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_windowHeightSelector, "-- Select column ...", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_matboardPositionSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_matboardClassSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_matboardTextureScaleSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_matboardTextureRotatSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_withGlassSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_glassPositionSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_glassClassSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_excelCriteriaSelector, "-- Select column ...", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_symbolFolderSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_classPictureClassSelector, "-- Manual", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaArtworkTitleSelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaAuthorNameSelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaArtworkMediaSelector, "-- Don't Import", 0)
        # vs.AddChoice(self.dialog, self.kWidgetID_metaTypeSelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaRoomLocationSelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaArtworkSourceSelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaRegistrationNumberSelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaDesignNotesSelector, "-- Don't Import", 0)
        vs.AddChoice(self.dialog, self.kWidgetID_metaExhibitionMediaSelector, "-- Don't Import", 0)

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
                                                self.kWidgetID_windowWidthSelector,
                                                self.parameters.windowWidthSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_windowWidthSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_windowHeightSelector,
                                                self.parameters.windowHeightSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_windowHeightSelector, selector_index, True)

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

        vs.SetBooleanItem(self.dialog, self.kWidgetID_symbolCreateSymbol,
                          self.parameters.symbolCreateSymbol == "True")

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_symbolFolderSelector,
                                                self.parameters.symbolFolderSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_symbolFolderSelector, selector_index, True)

        vs.SetBooleanItem(self.dialog, self.kWidgetID_classAssignPictureClass,
                          self.parameters.classAssignPictureClass == "True")

        selector_index = vs.GetPopUpChoiceIndex(self.dialog,
                                                self.kWidgetID_classPictureClassSelector,
                                                self.parameters.classClassPictureSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_classPictureClassSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_classPictureClass,
                                                self.parameters.pictureParameters.pictureClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_classPictureClass, selector_index, True)

        vs.SetBooleanItem(self.dialog, self.kWidgetID_classCreateMissingClasses,
                          self.parameters.createMissingClasses == "True")

        vs.SetBooleanItem(self.dialog, self.kWidgetID_metaImportMetadata,
                          self.parameters.metaImportMetadata == "True")

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaArtworkTitleSelector,
                                                self.parameters.metaArtworkTitleSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaArtworkTitleSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaAuthorNameSelector,
                                                self.parameters.metaAuthorNameSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaAuthorNameSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector,
                                                self.parameters.metaArtworkCreationDateSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaArtworkMediaSelector,
                                                self.parameters.metaArtworkMediaSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaArtworkMediaSelector, selector_index, True)

        # selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaTypeSelector,
        #                                         self.parameters.metaTypeSelector)
        # vs.SelectChoice(self.dialog, self.kWidgetID_metaTypeSelector, selector_index, True)
        #
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaRoomLocationSelector,
                                                self.parameters.metaRoomLocationSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaRoomLocationSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaArtworkSourceSelector,
                                                self.parameters.metaArtworkSourceSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaArtworkSourceSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaRegistrationNumberSelector,
                                                self.parameters.metaRegistrationNumberSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaRegistrationNumberSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector,
                                                self.parameters.metaAuthorBirthCountrySelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector,
                                                self.parameters.metaAuthorBirthDateSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector,
                                                self.parameters.metaAuthorDeathDateSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaDesignNotesSelector,
                                                self.parameters.metaDesignNotesSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaDesignNotesSelector, selector_index, True)

        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaExhibitionMediaSelector,
                                                self.parameters.metaExhibitionMediaSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaExhibitionMediaSelector, selector_index, True)

        self.update_criteria_values(True)

        manual = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_withImageSelector, 0) == 0
        enabled = vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage)
        vs.EnableItem(self.dialog, self.kWidgetID_withImage, manual)
        vs.EnableItem(self.dialog, self.kWidgetID_imageWidthLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_imageWidthSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_imageHeightLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_imageHeightSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_imagePositionLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_imagePositionSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_imagePosition, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_imageTextureLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_imageTextureSelector, not manual or enabled)

        manual = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_withFrameSelector, 0) == 0
        enabled = vs.GetBooleanItem(self.dialog, self.kWidgetID_withFrame)
        vs.EnableItem(self.dialog, self.kWidgetID_withFrame, manual)
        vs.EnableItem(self.dialog, self.kWidgetID_frameWidthLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameWidthSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameHeightLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameHeightSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameThicknessLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameThicknessSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameThickness, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameDepthLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameDepthSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameDepth, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameClassLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameClassSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameClass, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScaleLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScaleSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScale, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotationLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotationSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotation, not manual or enabled)

        manual = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_withMatboardSelector, 0) == 0
        enabled = vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard)
        vs.EnableItem(self.dialog, self.kWidgetID_withMatboard, manual)
        vs.EnableItem(self.dialog, self.kWidgetID_windowWidthLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_windowWidthSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_windowHeightLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_windowHeightSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardPositionLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardPositionSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardPosition, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardClassLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardClassSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardClass, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScaleLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScaleSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScale, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotatLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotatSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotat, not manual or enabled)

        manual = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_withGlassSelector, 0) == 0
        enabled = vs.GetBooleanItem(self.dialog, self.kWidgetID_withGlass)
        vs.EnableItem(self.dialog, self.kWidgetID_withGlass, manual)
        vs.EnableItem(self.dialog, self.kWidgetID_glassPositionLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_glassPositionSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_glassPosition, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_glassClassLabel, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_glassClassSelector, not manual or enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_glassClass, not manual or enabled)

        enabled = vs.GetBooleanItem(self.dialog, self.kWidgetID_symbolCreateSymbol)
        manual = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_symbolFolderSelector, 0) == 0
        vs.EnableItem(self.dialog, self.kWidgetID_symbolFolderSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_symbolFolder, enabled and manual)

        enabled = vs.GetBooleanItem(self.dialog, self.kWidgetID_classAssignPictureClass)
        manual = vs.GetSelectedChoiceIndex(self.dialog, self.kWidgetID_classPictureClassSelector, 0) == 0
        vs.EnableItem(self.dialog, self.kWidgetID_classPictureClassSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_classPictureClass, enabled and manual)

        enabled = vs.GetBooleanItem(self.dialog, self.kWidgetID_metaImportMetadata)
        vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkTitleLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkTitleSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorNameLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorNameSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkCreationDateLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkMediaLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkMediaSelector, enabled)
        # vs.EnableItem(self.dialog, self.kWidgetID_metaTypeLabel, enabled)
        # vs.EnableItem(self.dialog, self.kWidgetID_metaTypeSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaRoomLocationLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaRoomLocationSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkSourceLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaArtworkSourceSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaRegistrationNumberLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaRegistrationNumberSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorBirthCountryLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorBirthDateLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorDeathDateLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaDesignNotesLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaDesignNotesSelector, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaExhibitionMediaLabel, enabled)
        vs.EnableItem(self.dialog, self.kWidgetID_metaExhibitionMediaSelector, enabled)

        vs.SetBooleanItem(self.dialog, self.kWidgetID_importIgnoreErrors, self.parameters.importIgnoreErrors == "True")
        vs.SetBooleanItem(self.dialog, self.kWidgetID_importIgnoreExisting, self.parameters.importIgnoreExisting == "True")
        vs.SetBooleanItem(self.dialog, self.kWidgetID_importIgnoreUnmodified, self.parameters.importIgnoreUnmodified == "True")

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
        self.remove_field_options(self.kWidgetID_windowWidthSelector)
        self.remove_field_options(self.kWidgetID_windowHeightSelector)
        self.remove_field_options(self.kWidgetID_matboardPositionSelector)
        self.remove_field_options(self.kWidgetID_matboardClassSelector)
        self.remove_field_options(self.kWidgetID_matboardTextureScaleSelector)
        self.remove_field_options(self.kWidgetID_matboardTextureRotatSelector)
        self.remove_field_options(self.kWidgetID_withGlassSelector)
        self.remove_field_options(self.kWidgetID_glassPositionSelector)
        self.remove_field_options(self.kWidgetID_glassClassSelector)
        self.remove_field_options(self.kWidgetID_excelCriteriaSelector)
        self.remove_field_options(self.kWidgetID_excelCriteriaValue)
        self.remove_field_options(self.kWidgetID_symbolFolderSelector)
        self.remove_field_options(self.kWidgetID_classPictureClassSelector)
        self.remove_field_options(self.kWidgetID_metaArtworkTitleSelector)
        self.remove_field_options(self.kWidgetID_metaAuthorNameSelector)
        self.remove_field_options(self.kWidgetID_metaArtworkCreationDateSelector)
        self.remove_field_options(self.kWidgetID_metaArtworkMediaSelector)
        # self.remove_field_options(self.kWidgetID_metaTypeSelector)
        self.remove_field_options(self.kWidgetID_metaRoomLocationSelector)
        self.remove_field_options(self.kWidgetID_metaArtworkSourceSelector)
        self.remove_field_options(self.kWidgetID_metaRegistrationNumberSelector)
        self.remove_field_options(self.kWidgetID_metaAuthorBirthCountrySelector)
        self.remove_field_options(self.kWidgetID_metaAuthorBirthDateSelector)
        self.remove_field_options(self.kWidgetID_metaAuthorDeathDateSelector)
        self.remove_field_options(self.kWidgetID_metaDesignNotesSelector)
        self.remove_field_options(self.kWidgetID_metaExhibitionMediaSelector)

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

        # Window Width dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_windowWidthLabel, "Window Width: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_withMatboardLabel, self.kWidgetID_windowWidthLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_windowWidthSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_windowWidthSelector,
                                                self.parameters.windowWidthSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_windowWidthSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_windowWidthLabel, self.kWidgetID_windowWidthSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_windowWidthSelector, "Select the column for the matboard window width")
        # Window Height dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_windowHeightLabel, "Window Height: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_windowWidthLabel, self.kWidgetID_windowHeightLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_windowHeightSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_windowHeightSelector,
                                                self.parameters.windowHeightSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_windowHeightSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_windowHeightLabel, self.kWidgetID_windowHeightSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_windowHeightSelector, "Select the column for the matboard window height")

        # Matboard Position dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_matboardPositionLabel, "Matboard Position: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_windowHeightLabel, self.kWidgetID_matboardPositionLabel, 0, 0)
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
        vs.CreateGroupBox(self.dialog, self.kWidgetID_symbolGroup, "Symbol", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_excelCriteriaGroup, self.kWidgetID_symbolGroup, 0, 0)
        # Create Symbol checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_symbolCreateSymbol, "Create Symbol")
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_symbolGroup, self.kWidgetID_symbolCreateSymbol)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_symbolCreateSymbol, self.parameters.symbolCreateSymbol == "True")
        vs.SetHelpText(self.dialog, self.kWidgetID_symbolCreateSymbol, "Check to create a symbol for every Picture")
        # Symbol Folder
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_symbolFolderLabel, "Symbol Folder: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_symbolCreateSymbol, self.kWidgetID_symbolFolderLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_symbolFolderSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_symbolFolderSelector,
                                                self.parameters.symbolFolderSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_symbolFolderSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_symbolFolderLabel, self.kWidgetID_symbolFolderSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_symbolFolderSelector, "Select the column for the symbol folder name")

        vs.CreateEditText(self.dialog, self.kWidgetID_symbolFolder, self.parameters.symbolFolder, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_symbolFolderSelector, self.kWidgetID_symbolFolder, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_symbolFolder, "Enter the symbol folder name")

        # Class group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_classGroup, "Classes", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_symbolGroup, self.kWidgetID_classGroup, 0, 0)
        # Assign class checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_classAssignPictureClass, "Assign a Class to the pictures")
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_classGroup, self.kWidgetID_classAssignPictureClass)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_classAssignPictureClass, self.parameters.classAssignPictureClass == "True")
        vs.SetHelpText(self.dialog, self.kWidgetID_classAssignPictureClass, "Check to assign a Class to every Picture")
        # Picture Class
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_classPictureClassLabel, "Picture Class: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_classAssignPictureClass, self.kWidgetID_classPictureClassLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_classPictureClassSelector, input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_classPictureClassSelector,
                                                self.parameters.classClassPictureSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_classPictureClassSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_classPictureClassLabel, self.kWidgetID_classPictureClassSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_classPictureClassSelector, "Select the column for the picture class")
        vs.CreateClassPullDownMenu(self.dialog, self.kWidgetID_classPictureClass, input_field_width)
        class_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_classPictureClass, self.parameters.pictureParameters.pictureClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_classPictureClass, class_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_classPictureClassSelector, self.kWidgetID_classPictureClass, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_classPictureClass, "Enter the class of the picture here.")
        # Create missing classes
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_classCreateMissingClasses, "Create missing classes")
        vs.SetBelowItem(self.dialog, self.kWidgetID_classPictureClassLabel, self.kWidgetID_classCreateMissingClasses, 0, 0)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_classCreateMissingClasses, self.parameters.createMissingClasses == "True")
        vs.SetHelpText(self.dialog, self.kWidgetID_classCreateMissingClasses, "Create missing classes")

        # Metadata group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_metaGroup, "Metadata", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_classGroup, self.kWidgetID_metaGroup, 0, 0)

        # Import Metadata
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_metaImportMetadata, "Import Metada")
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_metaGroup, self.kWidgetID_metaImportMetadata)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_metaImportMetadata, self.parameters.metaImportMetadata == "True")
        vs.SetHelpText(self.dialog, self.kWidgetID_metaImportMetadata, "Select to import Artwork Metadata")
        # Picture Title
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaArtworkTitleLabel, "Artwork Title: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaImportMetadata, self.kWidgetID_metaArtworkTitleLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaArtworkTitleSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaArtworkTitleSelector,
                                                self.parameters.metaArtworkTitleSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaArtworkTitleSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaArtworkTitleLabel, self.kWidgetID_metaArtworkTitleSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaArtworkTitleSelector, "Select how Artwork Title has to be imported.")
        # Author Name
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaAuthorNameLabel, "Artwork Author: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaArtworkTitleLabel, self.kWidgetID_metaAuthorNameLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaAuthorNameSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaAuthorNameSelector,
                                                self.parameters.metaAuthorNameSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaAuthorNameSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaAuthorNameLabel, self.kWidgetID_metaAuthorNameSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaAuthorNameSelector, "Select how Artwork Author has to be imported.")
        # Creation Date
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaArtworkCreationDateLabel, "Creation Date: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaAuthorNameLabel, self.kWidgetID_metaArtworkCreationDateLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector,
                                                self.parameters.metaArtworkCreationDateSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaArtworkCreationDateLabel, self.kWidgetID_metaArtworkCreationDateSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaArtworkCreationDateSelector, "Select how Artwork Creation Date has to be imported.")
        # Artwork Media
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaArtworkMediaLabel, "Artwork Media: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaArtworkCreationDateLabel, self.kWidgetID_metaArtworkMediaLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaArtworkMediaSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaArtworkMediaSelector,
                                                self.parameters.metaArtworkMediaSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaArtworkMediaSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaArtworkMediaLabel, self.kWidgetID_metaArtworkMediaSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaArtworkMediaSelector, "Select how Artwork Media has to be imported.")
        # Room Location
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaRoomLocationLabel, "Room Location: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaArtworkMediaLabel, self.kWidgetID_metaRoomLocationLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaRoomLocationSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaRoomLocationSelector,
                                                self.parameters.metaRoomLocationSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaRoomLocationSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaRoomLocationLabel, self.kWidgetID_metaRoomLocationSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaRoomLocationSelector, "Select the field for the room location.")
        # Artwork Source
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaArtworkSourceLabel, "Artwork Source: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaRoomLocationLabel, self.kWidgetID_metaArtworkSourceLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaArtworkSourceSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaArtworkSourceSelector,
                                                self.parameters.metaArtworkSourceSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaArtworkSourceSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaArtworkSourceLabel, self.kWidgetID_metaArtworkSourceSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaArtworkSourceSelector, "Select how Artwork Source has to be imported.")
        # Registration Number
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaRegistrationNumberLabel, "Registration Number: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaArtworkSourceLabel, self.kWidgetID_metaRegistrationNumberLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaRegistrationNumberSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaRegistrationNumberSelector,
                                                self.parameters.metaRegistrationNumberSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaRegistrationNumberSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaRegistrationNumberLabel, self.kWidgetID_metaRegistrationNumberSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaRegistrationNumberSelector, "Select how Registration Number has to be imported.")
        # Author Birth Country
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaAuthorBirthCountryLabel, "Author Birth Country: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaRegistrationNumberLabel, self.kWidgetID_metaAuthorBirthCountryLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector,
                                                self.parameters.metaAuthorBirthCountrySelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaAuthorBirthCountryLabel, self.kWidgetID_metaAuthorBirthCountrySelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaAuthorBirthCountrySelector, "Select how Author Birth Country has to be imported.")
        # Author Birth Date
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaAuthorBirthDateLabel, "Author Birth Date: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaAuthorBirthCountryLabel, self.kWidgetID_metaAuthorBirthDateLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector,
                                                self.parameters.metaAuthorBirthDateSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaAuthorBirthDateLabel, self.kWidgetID_metaAuthorBirthDateSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaAuthorBirthDateSelector, "Select how Author Birth Date has to be imported.")
        # Author Death Date
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaAuthorDeathDateLabel, "Author Death Date: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaAuthorBirthDateLabel, self.kWidgetID_metaAuthorDeathDateLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector,
                                                self.parameters.metaAuthorDeathDateSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaAuthorDeathDateLabel, self.kWidgetID_metaAuthorDeathDateSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaAuthorDeathDateSelector, "Select how Author Death Date has to be imported.")
        # Design Notes
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaDesignNotesLabel, "Design Notes: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaAuthorDeathDateLabel, self.kWidgetID_metaDesignNotesLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaDesignNotesSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaDesignNotesSelector,
                                                self.parameters.metaDesignNotesSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaDesignNotesSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaDesignNotesLabel, self.kWidgetID_metaDesignNotesSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaDesignNotesSelector, "Select how Design Notes has to be imported.")
        # Exhibition Media
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_metaExhibitionMediaLabel, "Exhibition Media: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaDesignNotesLabel, self.kWidgetID_metaExhibitionMediaLabel, 0, 0)
        vs.CreatePullDownMenu(self.dialog, self.kWidgetID_metaExhibitionMediaSelector, 2 * input_field_width)
        selector_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_metaExhibitionMediaSelector,
                                                self.parameters.metaExhibitionMediaSelector)
        vs.SelectChoice(self.dialog, self.kWidgetID_metaExhibitionMediaSelector, selector_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_metaExhibitionMediaLabel, self.kWidgetID_metaExhibitionMediaSelector, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_metaExhibitionMediaSelector, "Select how Exhibition Media has to be imported.")

        # Import group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_importGroup, "Import", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_metaGroup, self.kWidgetID_importGroup, 0, 0)

        # Ignore Existing
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_importIgnoreExisting, "Ignore manual fields on existing Pictures")
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_importGroup, self.kWidgetID_importIgnoreExisting)
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

    def update_picture(self, picture_parameters: PictureParameters, existing_picture, log_file: IO):
        log_message = ""
        image_message = ""
        frame_message = ""
        matboard_message = ""
        glass_message = ""
        metadata_message = ""
        changed = False

        # existing_picture = vs.GetObject(picture_parameters.pictureName)
        if self.parameters.withImageSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
            if picture_parameters.withImage != vs.GetRField(existing_picture, "Picture", "WithImage"):
                if picture_parameters.withImage == "True":
                    image_message = "- Add image " + image_message
                else:
                    image_message = "- Removed image "
                vs.SetRField(existing_picture, "Picture", "WithImage", picture_parameters.withImage)
                changed = True

        if picture_parameters.withImage == "True":
            if not same_dimension(vs.GetRField(existing_picture, "Picture", "ImageWidth"), picture_parameters.imageWidth):
                image_message = image_message + "- Image With changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "ImageWidth"), picture_parameters.imageWidth))
                vs.SetRField(existing_picture, "Picture", "ImageWidth", picture_parameters.imageWidth)
                changed = True

            if not same_dimension(vs.GetRField(existing_picture, "Picture", "ImageHeight"), picture_parameters.imageHeight):
                image_message = image_message + "- Image Height changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "ImageHeight"), picture_parameters.imageHeight))
                vs.SetRField(existing_picture, "Picture", "ImageHeight", picture_parameters.imageHeight)
                changed = True

            if self.parameters.imagePositionSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "ImagePosition"), picture_parameters.imagePosition):
                    image_message = image_message + "- Image Position changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "ImagePosition"), picture_parameters.imagePosition))
                    vs.SetRField(existing_picture, "Picture", "ImagePosition", picture_parameters.imagePosition)
                    changed = True

        if self.parameters.withFrameSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
            if picture_parameters.withFrame != vs.GetRField(existing_picture, "Picture", "WithFrame"):
                if picture_parameters.withFrame == "True":
                    frame_message = "- Add frame " + frame_message
                else:
                    frame_message = "- Removed frame "
                vs.SetRField(existing_picture, "Picture", "WithFrame", picture_parameters.withFrame)
                changed = True

        if picture_parameters.withFrame == "True":
            if not same_dimension(vs.GetRField(existing_picture, "Picture", "FrameWidth"), picture_parameters.frameWidth):
                frame_message = frame_message + "- Frame Width changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "FrameWidth"), picture_parameters.frameWidth))
                vs.SetRField(existing_picture, "Picture", "FrameWidth", picture_parameters.frameWidth)
                changed = True

            if not same_dimension(vs.GetRField(existing_picture, "Picture", "FrameHeight"), picture_parameters.frameHeight):
                frame_message = frame_message + "- Frame Height changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "FrameHeight"), picture_parameters.frameHeight))
                vs.SetRField(existing_picture, "Picture", "FrameHeight", picture_parameters.frameHeight)
                changed = True

            if self.parameters.frameThicknessSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "FrameThickness"), picture_parameters.frameThickness):
                    frame_message = frame_message + "- Frame Thickness changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "FrameThickness"), picture_parameters.frameThickness))
                    vs.SetRField(existing_picture, "Picture", "FrameThickness", picture_parameters.frameThickness)
                    changed = True

            if self.parameters.frameDepthSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "FrameDepth"), picture_parameters.frameDepth):
                    frame_message = frame_message + "- Frame Depth changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "FrameDepth"), picture_parameters.frameDepth))
                    vs.SetRField(existing_picture, "Picture", "FrameDepth", picture_parameters.frameDepth)
                    changed = True

            if self.parameters.frameClassSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if picture_parameters.frameClass != vs.GetRField(existing_picture, "Picture", "FrameClass"):
                    frame_message = frame_message + "- Frame Class changed ({} --> {}) ".format(vs.GetRField(existing_picture, "Picture", "FrameClass"), picture_parameters.frameClass)
                    vs.SetRField(existing_picture, "Picture", "FrameClass", picture_parameters.frameClass)
                    changed = True

            if self.parameters.frameTextureScaleSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "FrameTextureScale"), picture_parameters.frameTextureScale):
                    frame_message = frame_message + "- Frame Texture Scale changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "FrameTextureScale"), picture_parameters.frameTextureScale))
                    vs.SetRField(existing_picture, "Picture", "FrameTextureScale", picture_parameters.frameTextureScale)
                    changed = True

            if self.parameters.frameTextureRotationSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "FrameTextureRotation"), picture_parameters.frameTextureRotation):
                    frame_message = frame_message + "- Frame Texture Rotation changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "FrameTextureRotation"), picture_parameters.frameTextureRotation))
                    vs.SetRField(existing_picture, "Picture", "FrameTextureRotation", picture_parameters.frameTextureRotation)
                    changed = True

        if self.parameters.withMatboardSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
            if picture_parameters.withMatboard != vs.GetRField(existing_picture, "Picture", "WithMatboard"):
                if picture_parameters.withMatboard == "True":
                    matboard_message = "- Add matboard " + matboard_message
                else:
                    matboard_message = "- Removed matboard "
                vs.SetRField(existing_picture, "Picture", "WithMatboard", picture_parameters.withMatboard)
                changed = True

        if picture_parameters.withMatboard == "True":
            if self.parameters.frameWidthSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "FrameWidth"), picture_parameters.frameWidth):
                    matboard_message = matboard_message + "- Frame Width changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "FrameWidth"), picture_parameters.frameWidth))
                    vs.SetRField(existing_picture, "Picture", "FrameWidth", picture_parameters.frameWidth)
                    changed = True

            if self.parameters.frameHeightSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "FrameHeight"), picture_parameters.frameHeight):
                    matboard_message = matboard_message + "- Frame Height changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "FrameHeight"), picture_parameters.frameHeight))
                    vs.SetRField(existing_picture, "Picture", "FrameHeight", picture_parameters.frameHeight)
                    changed = True

            if self.parameters.windowWidthSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "WindowWidth"), picture_parameters.windowWidth):
                    matboard_message = matboard_message + "- Matboard window width changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "WindowWidth"), picture_parameters.windowWidth))
                    vs.SetRField(existing_picture, "Picture", "WindowWidth", picture_parameters.windowWidth)
                    changed = True

            if self.parameters.windowHeightSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "WindowHeight"), picture_parameters.windowHeight):
                    matboard_message = matboard_message + "- Matboard window height changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "WindowHeight"), picture_parameters.windowHeight))
                    vs.SetRField(existing_picture, "Picture", "WindowHeight", picture_parameters.windowHeight)
                    changed = True

            if self.parameters.matboardPositionSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "MatboardPosition"), picture_parameters.matboardPosition):
                    matboard_message = matboard_message + "- Matboard Position changed ({0[0]} --> {0[1]}) ".format(
                        dimension_strings(vs.GetRField(existing_picture, "Picture", "MatboardPosition"), picture_parameters.matboardPosition))

            if self.parameters.matboardClassSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if picture_parameters.matboardClass != vs.GetRField(existing_picture, "Picture", "MatboardClass"):
                    matboard_message = matboard_message + "- Matboard Class changed ({} --> {}) ".format(vs.GetRField(existing_picture, "Picture", "MatboardClass"), picture_parameters.matboardClass)
                    vs.SetRField(existing_picture, "Picture", "MatboardClass", picture_parameters.matboardClass)
                    changed = True

            if self.parameters.matboardTextureScaleSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "MatboardTextureScale"), picture_parameters.matboardTextureScale):
                    matboard_message = matboard_message + "- Matboard Texture Scale changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "MatboardTextureScale"), picture_parameters.matboardTextureScale))
                    vs.SetRField(existing_picture, "Picture", "MatboardTextureScale", picture_parameters.matboardTextureScale)
                    changed = True

            if self.parameters.matboardTextureRotatSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "MatboardTextureRotat"), picture_parameters.matboardTextureRotat):
                    matboard_message = matboard_message + "- Matboard Texture Rotation changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "MatboardTextureRotat"), picture_parameters.matboardTextureRotat))
                    vs.SetRField(existing_picture, "Picture", "MatboardTextureRotat", picture_parameters.matboardTextureRotat)
                    changed = True

        if self.parameters.withGlassSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
            if picture_parameters.withGlass != vs.GetRField(existing_picture, "Picture", "WithGlass"):
                if picture_parameters.withGlass == "True":
                    glass_message = "- Add glass " + image_message
                else:
                    glass_message = "- Removed glass "
                vs.SetRField(existing_picture, "Picture", "WithGlass", picture_parameters.withGlass)
                changed = True

        if picture_parameters.withGlass == "True":
            if self.parameters.glassPositionSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if not same_dimension(vs.GetRField(existing_picture, "Picture", "GlassPosition"), picture_parameters.glassPosition):
                    glass_message = glass_message + "- Glass Position changed ({0[0]} --> {0[1]}) ".format(dimension_strings(vs.GetRField(existing_picture, "Picture", "GlassPosition"), picture_parameters.glassPosition))
                    vs.SetRField(existing_picture, "Picture", "GlassPosition", picture_parameters.glassPosition)
                    changed = True

            if self.parameters.glassClassSelector != "-- Manual" or self.parameters.importIgnoreExisting == "False":
                if picture_parameters.glassClass != vs.GetRField(existing_picture, "Picture", "GlassClass"):
                    glass_message = glass_message + "- Glass Class changed ({} --> {}) ".format(vs.GetRField(existing_picture, "Picture", "GlassClass"), picture_parameters.glassClass)
                    vs.SetRField(existing_picture, "Picture", "GlassClass", picture_parameters.glassClass)
                    changed = True

        # Update Metadata information
        existing_symbol = vs.GetObject("{} Picture Symbol".format(vs.GetName(existing_picture)))
        if self.parameters.metaImportMetadata == "True" and self.parameters.importIgnoreExisting == "False" and existing_symbol:

            if picture_parameters.withImage == "True":
                vs.SetRField(existing_symbol, "Object list data", "Image size", "Height: {}, Width: {}".format(picture_parameters.imageHeight, picture_parameters.imageWidth))
            if picture_parameters.withFrame == "True" or picture_parameters.withMatboard == "True":
                vs.SetRField(existing_symbol, "Object list data", "Frame size", "Height: {}, Width: {}".format(picture_parameters.frameHeight, picture_parameters.frameWidth))
            if picture_parameters.withMatboard == "True":
                vs.SetRField(existing_symbol, "Object list data", "Window size", "Height: {}, Width: {}".format(picture_parameters.windowHeight, picture_parameters.windowWidth))

            if self.parameters.metaArtworkTitleSelector != "-- Don't Import":
                if self.parameters.pictureRecord.artworkTitle is None:
                    self.parameters.pictureRecord.artworkTitle = ""
                if self.parameters.pictureRecord.artworkTitle != vs.GetRField(existing_symbol, "Object list data", "Artwork title"):
                    metadata_message += "Artwork Title changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Artwork title"), self.parameters.pictureRecord.artworkTitle)
                    vs.SetRField(existing_symbol, "Object list data", "Artwork title", self.parameters.pictureRecord.artworkTitle)
                    changed = True

            if self.parameters.metaAuthorNameSelector != "-- Don't Import":
                if self.parameters.pictureRecord.authorName is None:
                    self.parameters.pictureRecord.authorName = ""
                if self.parameters.pictureRecord.authorName != vs.GetRField(existing_symbol, "Object list data", "Author name"):
                    metadata_message += "Author name changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Author name"), self.parameters.pictureRecord.authorName)
                    vs.SetRField(existing_symbol, "Object list data", "Author name", self.parameters.pictureRecord.authorName)
                    changed = True

            if self.parameters.metaArtworkCreationDateSelector != "-- Don't Import":
                if self.parameters.pictureRecord.artworkCreationDate is None:
                    self.parameters.pictureRecord.artworkCreationDate = ""
                if self.parameters.pictureRecord.artworkCreationDate != vs.GetRField(existing_symbol, "Object list data", "Artwork creation date"):
                    metadata_message += "Artwork creation date changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Artwork creation date"), self.parameters.pictureRecord.artworkCreationDate)
                    vs.SetRField(existing_symbol, "Object list data", "Artwork creation date", self.parameters.pictureRecord.artworkCreationDate)
                    changed = True

            if self.parameters.metaArtworkMediaSelector != "-- Don't Import":
                if self.parameters.pictureRecord.artworkMedia is None:
                    self.parameters.pictureRecord.artworkMedia = ""
                if self.parameters.pictureRecord.artworkMedia != vs.GetRField(existing_symbol, "Object list data", "Artwork media"):
                    metadata_message += "Artwork media changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Artwork media"), self.parameters.pictureRecord.artworkMedia)
                    vs.SetRField(existing_symbol, "Object list data", "Artwork media", self.parameters.pictureRecord.artworkMedia)
                    changed = True

            # if self.settings.metaTypeSelector != "-- Don't Import":
            #     self.settings.pictureRecord. = row[self.settings.metaTypeSelector.lower()]

            if self.parameters.metaRoomLocationSelector != "-- Don't Import":
                if self.parameters.pictureRecord.roomLocation is None:
                    self.parameters.pictureRecord.roomLocation = ""
                if self.parameters.pictureRecord.roomLocation != vs.GetRField(existing_symbol, "Object list data", "Room Location"):
                    metadata_message += "Room Location changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Room Location"), self.parameters.pictureRecord.roomLocation)
                    vs.SetRField(existing_symbol, "Object list data", "Room Location", self.parameters.pictureRecord.roomLocation)
                    changed = True

            if self.parameters.metaArtworkSourceSelector != "-- Don't Import":
                if self.parameters.pictureRecord.artworkSource is None:
                    self.parameters.pictureRecord.artworkSource = ""
                if self.parameters.pictureRecord.artworkSource != vs.GetRField(existing_symbol, "Object list data", "Artwork source/lender"):
                    metadata_message += "Artwork source/lender changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Artwork source/lender"), self.parameters.pictureRecord.artworkSource)
                    vs.SetRField(existing_symbol, "Object list data", "Artwork source/lender", self.parameters.pictureRecord.artworkSource)
                    changed = True

            if self.parameters.metaRegistrationNumberSelector != "-- Don't Import":
                if self.parameters.pictureRecord.registrationNumber is None:
                    self.parameters.pictureRecord.registrationNumber = ""
                if self.parameters.pictureRecord.registrationNumber != vs.GetRField(existing_symbol, "Object list data", "WDFM registration number"):
                    metadata_message += "WDFM registration number changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "WDFM registration number"), self.parameters.pictureRecord.registrationNumber)
                    vs.SetRField(existing_symbol, "Object list data", "WDFM registration number", self.parameters.pictureRecord.registrationNumber)
                    changed = True

            if self.parameters.metaAuthorBirthCountrySelector != "-- Don't Import":
                if self.parameters.pictureRecord.authorBirthCountry is None:
                    self.parameters.pictureRecord.authorBirthCountry = ""
                if self.parameters.pictureRecord.authorBirthCountry != vs.GetRField(existing_symbol, "Object list data", "Author birth country"):
                    metadata_message += "Author birth country changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Author birth country"), self.parameters.pictureRecord.authorBirthCountry)
                    vs.SetRField(existing_symbol, "Object list data", "Author birth country", self.parameters.pictureRecord.authorBirthCountry)
                    changed = True

            if self.parameters.metaAuthorBirthDateSelector != "-- Don't Import":
                if self.parameters.pictureRecord.authorBirthDate is None:
                    self.parameters.pictureRecord.authorBirthDate = ""
                if self.parameters.pictureRecord.authorBirthDate != vs.GetRField(existing_symbol, "Object list data", "Author date of birth"):
                    metadata_message += "Author date of birth changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Author date of birth"), self.parameters.pictureRecord.authorBirthDate)
                    vs.SetRField(existing_symbol, "Object list data", "Author date of birth", self.parameters.pictureRecord.authorBirthDate)
                    changed = True

            if self.parameters.metaAuthorDeathDateSelector != "-- Don't Import":
                if self.parameters.pictureRecord.authorDeathDate is None:
                    self.parameters.pictureRecord.authorDeathDate = ""
                if self.parameters.pictureRecord.authorDeathDate != vs.GetRField(existing_symbol, "Object list data", "Author date of death"):
                    metadata_message += "Author date of death changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Author date of death"), self.parameters.pictureRecord.authorDeathDate)
                    vs.SetRField(existing_symbol, "Object list data", "Author date of death", self.parameters.pictureRecord.authorDeathDate)
                    changed = True

            if self.parameters.metaDesignNotesSelector != "-- Don't Import":
                if self.parameters.pictureRecord.designNotes is None:
                    self.parameters.pictureRecord.designNotes = ""
                if self.parameters.pictureRecord.designNotes != vs.GetRField(existing_symbol, "Object list data", "Design notes"):
                    metadata_message += "Design notes changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Design notes"), self.parameters.pictureRecord.designNotes)
                    vs.SetRField(existing_symbol, "Object list data", "Design notes", self.parameters.pictureRecord.designNotes)
                    changed = True

            if self.parameters.metaExhibitionMediaSelector != "-- Don't Import":
                if self.parameters.pictureRecord.exhibitionMedia is None:
                    self.parameters.pictureRecord.exhibitionMedia = ""
                if self.parameters.pictureRecord.exhibitionMedia != vs.GetRField(existing_symbol, "Object list data", "Exhibition media"):
                    metadata_message += "Exhibition media changed ({} --> {}) ".format(vs.GetRField(existing_symbol, "Object list data", "Exhibition media"), self.parameters.pictureRecord.exhibitionMedia)
                    vs.SetRField(existing_symbol, "Object list data", "Exhibition media", self.parameters.pictureRecord.exhibitionMedia)
                    changed = True

        if changed:
            vs.ResetObject(existing_picture)

            log_message = "{} * [Modified] ".format(picture_parameters.pictureName) + image_message + frame_message + matboard_message + glass_message + metadata_message + "\n"
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
                            texture_name = "Arroway {}".format(picture_parameters.pictureName.replace('-', ' ').replace('_', ' ')) \
                                           + ' ' + str(inner) + ' ' + str(outer)
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
                if not self.parameters.classAssignPictureClass:
                    picture_parameters.pictureClass = ""

                picture_parameters.imageTexture = vs.GetName(texture)
                build_picture(picture_parameters, self.parameters.pictureRecord if self.parameters.metaImportMetadata == "True" else None)

                log_message = "{} * [New] \n".format(picture_parameters.pictureName)
                self.importNewCount += 1

            log_file.write(log_message)

    def import_pictures(self):
        self.importNewCount = 0
        self.importUpdatedCount = 0
        self.importDeletedCount = 0
        self.importErrorCount = 0
        document_file_name = vs.GetFPathName()
        document_folder = os.path.dirname(document_file_name)
        if not document_folder:
            # document_folder = "C:/tmp"
            vs.AlertCritical("Save the file first", "Before importing Pictures you must name and save the document")
            self.result = kCancel
            return

        log_file_name = document_folder + "/" + "Import_Pictures_" + strftime("%y_%m_%d_%H_%M_%S", gmtime()) + ".log"

        # log_file = open(log_file_name, "w")
        with open(log_file_name, "w") as log_file:
            try:
                vs.ProgressDlgOpen("Importing Pictures", True)
                total_rows = self.excel.get_worksheet_row_count()
                vs.ProgressDlgSetMeter("Importing " + str(total_rows) + " Pictures ...")
                vs.ProgressDlgStart(100.0, total_rows)

                for picture_parameters in self.excel.get_worksheet_rows(log_file):

                    if vs.ProgressDlgHasCancel():
                        break
                    vs.ProgressDlgYield(1)
                    vs.ProgressDlgSetTopMsg("New Pictures: {}".format(self.importNewCount))
                    vs.ProgressDlgSetBotMsg("Modified Pictures: {}".format(self.importUpdatedCount))

                    if picture_parameters.pictureName:
                        existing_picture = vs.GetObject(picture_parameters.pictureName)
                        if not existing_picture:
                            existing_symbol = vs.GetObject("{} Picture Symbol".format(picture_parameters.pictureName))
                            if existing_symbol:
                                if vs.GetTypeN(existing_symbol) == 16:
                                    existing_picture = vs.FInSymDef(existing_symbol)
                                    while existing_picture:
                                        if vs.GetTypeN(existing_picture) == 86:
                                            break
                                        existing_picture = vs.NextObj(existing_picture)

                        if existing_picture:
                            self.update_picture(picture_parameters, existing_picture, log_file)
                        else:
                            self.new_picture(picture_parameters, log_file)
                    else:
                        self.importErrorCount += 1
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
            except IOError as e:
                vs.AlertCritical("Cannot open log file", "I/O error({0}): {1}".format(e.errno, e.strerror))

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
