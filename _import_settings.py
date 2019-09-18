import vs
from _picture_settings import PictureParameters


class ImportSettings:
    def __init__(self):
        self.pictureParameters = PictureParameters()
        # self.errorString = ""
        # Picture parameters
        valid, self.pictureParameters.withImage = vs.GetSavedSetting("importPictures", "withImage")
        valid, value = vs.GetSavedSetting("importPictures", "imageWidth")
        if valid:
            valid, self.pictureParameters.imageWidth = vs.ValidNumStr(value)
        valid, value = vs.GetSavedSetting("importPictures", "imageHeight")
        if valid:
            valid, self.pictureParameters.imageHeight = vs.ValidNumStr(value)
        valid, value = vs.GetSavedSetting("importPictures", "imagePosition")
        if valid:
            valid, self.pictureParameters.imagePosition = vs.ValidNumStr(value)
        valid, self.pictureParameters.imageTexutre = vs.GetSavedSetting("importPictures", "imageTexutre")
        valid, self.pictureParameters.withFrame = vs.GetSavedSetting("importPictures", "withFrame")
        valid, value = vs.GetSavedSetting("importPictures", "frameWidth")
        if valid:
            valid, self.pictureParameters.frameWidth = vs.ValidNumStr(value)
        valid, value = vs.GetSavedSetting("importPictures", "frameHeight")
        if valid:
            valid, self.pictureParameters.frameHeight = vs.ValidNumStr(value)
        valid, value = vs.GetSavedSetting("importPictures", "frameThickness")
        if valid:
            valid, self.pictureParameters.frameThickness = vs.ValidNumStr(value)
        valid, value = vs.GetSavedSetting("importPictures", "frameDepth")
        if valid:
            valid, self.pictureParameters.frameDepth = vs.ValidNumStr(value)
        valid, self.pictureParameters.frameClass = vs.GetSavedSetting("importPictures", "frameClass")
        valid, value = vs.GetSavedSetting("importPictures", "frameTextureScale")
        if valid:
            valid, self.pictureParameters.frameTextureScale = vs.ValidNumStr(value)
        valid, value = vs.GetSavedSetting("importPictures", "frameTextureRotation")
        if valid:
            valid, self.pictureParameters.frameTextureRotation = vs.ValidNumStr(value)
        valid, self.pictureParameters.withMatboard = vs.GetSavedSetting("importPictures", "withMatboard")
        valid, value = vs.GetSavedSetting("importPictures", "matboardPosition")
        if valid:
            valid, self.pictureParameters.matboardPosition = vs.ValidNumStr(value)
        valid, self.pictureParameters.matboardClass = vs.GetSavedSetting("importPictures", "matboardClass")
        valid, value = vs.GetSavedSetting("importPictures", "matboardTextureScale")
        if valid:
            valid, self.pictureParameters.matboardTextureScale = vs.ValidNumStr(value)
        valid, value = vs.GetSavedSetting("importPictures", "matboardTextureRotat")
        if valid:
            valid, self.pictureParameters.matboardTextureRotat = vs.ValidNumStr(value)
        valid, self.pictureParameters.withGlass = vs.GetSavedSetting("importPictures", "withGlass")
        valid, value = vs.GetSavedSetting("importPictures", "glassPosition")
        if valid:
            valid, self.pictureParameters.glassPosition = vs.ValidNumStr(value)
        valid, self.pictureParameters.glassClass = vs.GetSavedSetting("importPictures", "glassClass")

        # valid, self.pictureParameters.withImage = vs.GetSavedSetting("importPictures", "withImage")
        # if not valid:
        #     self.pictureParameters.withImage = "True"
        # valid, value = vs.GetSavedSetting("importPictures", "imageWidth")
        # if valid:
        #     valid, self.pictureParameters.imageWidth = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.imageWidth = 10.0
        # valid, value = vs.GetSavedSetting("importPictures", "imageHeight")
        # if valid:
        #     valid, self.pictureParameters.imageHeight = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.imageHeight = 6.0
        # valid, value = vs.GetSavedSetting("importPictures", "imagePosition")
        # if valid:
        #     valid, self.pictureParameters.imagePosition = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.imagePosition = 0.3
        # valid, self.pictureParameters.imageTexutre = vs.GetSavedSetting("importPictures", "imageTexutre")
        # if not valid:
        #     self.pictureParameters.imageTexutre = ""
        # valid, self.pictureParameters.withFrame = vs.GetSavedSetting("importPictures", "withFrame")
        # if not valid:
        #     self.pictureParameters.withFrame = "True"
        # valid, value = vs.GetSavedSetting("importPictures", "frameWidth")
        # if valid:
        #     valid, self.pictureParameters.frameWidth = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.frameWidth = 8.0
        # valid, value = vs.GetSavedSetting("importPictures", "frameHeight")
        # if valid:
        #     valid, self.pictureParameters.frameHeight = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.frameHeight = 12.0
        # valid, value = vs.GetSavedSetting("importPictures", "frameThickness")
        # if valid:
        #     valid, self.pictureParameters.frameThickness = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.frameThickness = 1.0
        # valid, value = vs.GetSavedSetting("importPictures", "frameDepth")
        # if valid:
        #     valid, self.pictureParameters.frameDepth = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.frameDepth = 1.0
        # valid, self.pictureParameters.frameClass = vs.GetSavedSetting("importPictures", "frameClass")
        # if not valid:
        #     self.pictureParameters.frameClass = "None"
        # valid, value = vs.GetSavedSetting("importPictures", "frameTextureScale")
        # if valid:
        #     valid, self.pictureParameters.frameTextureScale = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.frameTextureScale = 0.1
        # valid, value = vs.GetSavedSetting("importPictures", "frameTextureRotation")
        # if valid:
        #     valid, self.pictureParameters.frameTextureRotation = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.frameTextureRotation = 0.0
        # valid, self.pictureParameters.withMatboard = vs.GetSavedSetting("importPictures", "withMatboard")
        # if not valid:
        #     self.pictureParameters.withMatboard = "True"
        # valid, value = vs.GetSavedSetting("importPictures", "matboardPosition")
        # if valid:
        #     valid, self.pictureParameters.matboardPosition = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.matboardPosition = 0.25
        # valid, self.pictureParameters.matboardClass = vs.GetSavedSetting("importPictures", "matboardClass")
        # if not valid:
        #     self.pictureParameters.matboardClass = "None"
        # valid, value = vs.GetSavedSetting("importPictures", "matboardTextureScale")
        # if valid:
        #     valid, self.pictureParameters.matboardTextureScale = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.matboardTextureScale = 0.1
        # valid, value = vs.GetSavedSetting("importPictures", "matboardTextureRotat")
        # if valid:
        #     valid, self.pictureParameters.matboardTextureRotat = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.matboardTextureRotat = 0.0
        # valid, self.withGlass = vs.GetSavedSetting("importPictures", "withGlass")
        # if not valid:
        #     self.pictureParameters.withGlass = "True"
        # valid, value = vs.GetSavedSetting("importPictures", "glassPosition")
        # if valid:
        #     valid, self.pictureParameters.glassPosition = vs.ValidNumStr(value)
        # if not valid:
        #     self.pictureParameters.glassPosition = 0.75
        # valid, self.pictureParameters.glassClass = vs.GetSavedSetting("importPictures", "glassClass")
        # if not valid:
        #     self.pictureParameters.glassClass = "None"

        # valid, self.imageFolderName = vs.GetSavedSetting("importPictures", "imageFolderName")
        # if not valid:
        #     self.imageFolderName = "Enter images folder name"
        # Dialog settings
        valid, self.excelFileName = vs.GetSavedSetting("importPictures", "excelFileName")
        if not valid:
            self.excelFileName = "Enter excel file name"
        valid, self.excelSheetName = vs.GetSavedSetting("importPictures", "excelSheetName")
        if not valid:
            self.excelSheetName = "Select an excel sheet"
        valid, self.withImageSelector = vs.GetSavedSetting("importPictures", "withImageSelector")
        if not valid:
            self.withImageSelector = "-- Manual"
        valid, self.imageWidthSelector = vs.GetSavedSetting("importPictures", "imageWidthSelector")
        if not valid:
            self.imageWidthSelector = "-- Select column ..."
        valid, self.imageHeightSelector = vs.GetSavedSetting("importPictures", "imageHeightSelector")
        if not valid:
            self.imageHeightSelector = "-- Select column ..."
        valid, self.imagePositionSelector = vs.GetSavedSetting("importPictures", "imagePositionSelector")
        if not valid:
            self.imagePositionSelector = "-- Manual"
        valid, self.imageTextureSelector = vs.GetSavedSetting("importPictures", "imageTextureSelector")
        if not valid:
            self.imageTextureSelector = "-- Select column ..."
        valid, self.withFrameSelector = vs.GetSavedSetting("importPictures", "withFrameSelector")
        if not valid:
            self.withFrameSelector = "-- Manual"
        valid, self.frameWidthSelector = vs.GetSavedSetting("importPictures", "frameWidthSelector")
        if not valid:
            self.frameWidthSelector = "-- Select column ..."
        valid, self.frameHeightSelector = vs.GetSavedSetting("importPictures", "frameHeightSelector")
        if not valid:
            self.frameHeightSelector = "-- Select column ..."
        valid, self.frameThicknessSelector = vs.GetSavedSetting("importPictures", "frameThicknessSelector")
        if not valid:
            self.frameThicknessSelector = "-- Manual"
        valid, self.frameDepthSelector = vs.GetSavedSetting("importPictures", "frameDepthSelector")
        if not valid:
            self.frameDepthSelector = "-- Manual"
        valid, self.frameClassSelector = vs.GetSavedSetting("importPictures", "frameClassSelector")
        if not valid:
            self.frameClassSelector = "-- Manual"
        valid, self.frameTextureScaleSelector = vs.GetSavedSetting("importPictures", "frameTextureScaleSelector")
        if not valid:
            self.frameTextureScaleSelector = "-- Manual"
        valid, self.frameTextureRotationSelector = vs.GetSavedSetting("importPictures", "frameTextureRotationSelector")
        if not valid:
            self.frameTextureRotationSelector = "-- Manual"
        valid, self.withMatboardSelector = vs.GetSavedSetting("importPictures", "withMatboardSelector")
        if not valid:
            self.withMatboardSelector = "-- Manual"
        valid, self.matboardPositionSelector = vs.GetSavedSetting("importPictures", "matboardPositionSelector")
        if not valid:
            self.matboardPositionSelector = "-- Manual"
        valid, self.matboardClassSelector = vs.GetSavedSetting("importPictures", "matboardClassSelector")
        if not valid:
            self.matboardClassSelector = "-- Manual"
        valid, self.matboardTextureScaleSelector = vs.GetSavedSetting("importPictures", "matboardTextureScaleSelector")
        if not valid:
            self.matboardTextureScaleSelector = "-- Manual"
        valid, self.matboardTextureRotatSelector = vs.GetSavedSetting("importPictures", "matboardTextureRotatSelector")
        if not valid:
            self.matboardTextureRotatSelector = "-- Manual"
        valid, self.withGlassSelector = vs.GetSavedSetting("importPictures", "withGlassSelector")
        if not valid:
            self.withGlassSelector = "-- Manual"
        valid, self.glassPositionSelector = vs.GetSavedSetting("importPictures", "glassPositionSelector")
        if not valid:
            self.glassPositionSelector = "-- Manual"
        valid, self.glassClassSelector = vs.GetSavedSetting("importPictures", "glassClassSelector")
        if not valid:
            self.glassClassSelector = "-- Manual"
        valid, self.excelCriteriaSelector = vs.GetSavedSetting("importPictures", "excelCriteriaSelector")
        if not valid:
            self.excelCriteriaSelector = "-- Select column ..."
        valid, self.excelCriteriaValue = vs.GetSavedSetting("importPictures", "excelCriteriaValue")
        if not valid:
            self.excelCriteriaValue = "-- Select a value ..."
        valid, self.symbolCreateSymbol = vs.GetSavedSetting("importPictures", "symbolCreateSymbol")
        if not valid:
            self.symbolCreateSymbol = "True"
        valid, self.symbolFolderSelector = vs.GetSavedSetting("importPictures", "symbolFolderSelector")
        if not valid:
            self.symbolFolderSelector = "-- Manual"
        valid, self.symbolFolder = vs.GetSavedSetting("importPictures", "symbolFolder")
        if not valid:
            self.symbolFolder = "Pictures"
        valid, self.importIgnoreErrors = vs.GetSavedSetting("importPictures", "importIgnoreErrors")
        if not valid:
            self.importIgnoreErrors = "False"
        valid, self.importIgnoreExisting = vs.GetSavedSetting("importPictures", "importIgnoreExisting")
        if not valid:
            self.importIgnoreExisting = "False"
        valid, self.importIgnoreUnmodified = vs.GetSavedSetting("importPictures", "importIgnoreUnmodified")
        if not valid:
            self.importIgnoreUnmodified = "False"

    def save(self):

        #Picture parameters
        vs.SetSavedSetting("importPictures", "withImage", self.pictureParameters.withImage)
        vs.SetSavedSetting("importPictures", "imageWidth", str(self.pictureParameters.imageWidth))
        vs.SetSavedSetting("importPictures", "imageHeight", str(self.pictureParameters.imageHeight))
        vs.SetSavedSetting("importPictures", "imagePosition", str(self.pictureParameters.imagePosition))
        vs.SetSavedSetting("importPictures", "imageTexutre", self.pictureParameters.imageTexutre)
        vs.SetSavedSetting("importPictures", "withFrame", self.pictureParameters.withFrame)
        vs.SetSavedSetting("importPictures", "frameWidth", str(self.pictureParameters.frameWidth))
        vs.SetSavedSetting("importPictures", "frameHeight", str(self.pictureParameters.frameHeight))
        vs.SetSavedSetting("importPictures", "frameThickness", str(self.pictureParameters.frameThickness))
        vs.SetSavedSetting("importPictures", "frameDepth", str(self.pictureParameters.frameDepth))
        vs.SetSavedSetting("importPictures", "frameClass", self.pictureParameters.frameClass)
        vs.SetSavedSetting("importPictures", "frameTextureScale", str(self.pictureParameters.frameTextureScale))
        vs.SetSavedSetting("importPictures", "frameTextureRotation", str(self.pictureParameters.frameTextureRotation))
        vs.SetSavedSetting("importPictures", "withMatboard", self.pictureParameters.withMatboard)
        vs.SetSavedSetting("importPictures", "matboardPosition", str(self.pictureParameters.matboardPosition))
        vs.SetSavedSetting("importPictures", "matboardClass", self.pictureParameters.matboardClass)
        vs.SetSavedSetting("importPictures", "matboardTextureScale", str(self.pictureParameters.matboardTextureScale))
        vs.SetSavedSetting("importPictures", "matboardTextureRotat", str(self.pictureParameters.matboardTextureRotat))
        vs.SetSavedSetting("importPictures", "withGlass", self.pictureParameters.withGlass)
        vs.SetSavedSetting("importPictures", "glassPosition", str(self.pictureParameters.glassPosition))
        vs.SetSavedSetting("importPictures", "glassClass", self.pictureParameters.glassClass)

        # Dialog settings
        vs.SetSavedSetting("importPictures", "excelFileName", self.excelFileName)
        vs.SetSavedSetting("importPictures", "excelSheetName", self.excelSheetName)
        # vs.SetSavedSetting("importPictures", "imageFolderName", self.imageFolderName)
        vs.SetSavedSetting("importPictures", "withImageSelector", self.withImageSelector)
        vs.SetSavedSetting("importPictures", "imageWidthSelector", self.imageWidthSelector)
        vs.SetSavedSetting("importPictures", "imageHeightSelector", self.imageHeightSelector)
        vs.SetSavedSetting("importPictures", "imagePositionSelector", self.imagePositionSelector)
        vs.SetSavedSetting("importPictures", "imageTextureSelector", self.imageTextureSelector)
        vs.SetSavedSetting("importPictures", "withFrameSelector", self.withFrameSelector)
        vs.SetSavedSetting("importPictures", "frameWidthSelector", self.frameWidthSelector)
        vs.SetSavedSetting("importPictures", "frameHeightSelector", self.frameHeightSelector)
        vs.SetSavedSetting("importPictures", "frameThicknessSelector", self.frameThicknessSelector)
        vs.SetSavedSetting("importPictures", "frameDepthSelector", self.frameDepthSelector)
        vs.SetSavedSetting("importPictures", "frameClassSelector", self.frameClassSelector)
        vs.SetSavedSetting("importPictures", "frameTextureScaleSelector", self.frameTextureScaleSelector)
        vs.SetSavedSetting("importPictures", "frameTextureRotationSelector", self.frameTextureRotationSelector)
        vs.SetSavedSetting("importPictures", "withMatboardSelector", self.withMatboardSelector)
        vs.SetSavedSetting("importPictures", "matboardPositionSelector", self.matboardPositionSelector)
        vs.SetSavedSetting("importPictures", "matboardClassSelector", self.matboardClassSelector)
        vs.SetSavedSetting("importPictures", "matboardTextureScaleSelector", self.matboardTextureScaleSelector)
        vs.SetSavedSetting("importPictures", "matboardTextureRotatSelector", self.matboardTextureRotatSelector)
        vs.SetSavedSetting("importPictures", "withGlassSelector", self.withGlassSelector)
        vs.SetSavedSetting("importPictures", "glassPositionSelector", self.glassPositionSelector)
        vs.SetSavedSetting("importPictures", "glassClassSelector", self.glassClassSelector)
        vs.SetSavedSetting("importPictures", "excelCriteriaSelector", self.excelCriteriaSelector)
        vs.SetSavedSetting("importPictures", "excelCriteriaValue", self.excelCriteriaValue)
        vs.SetSavedSetting("importPictures", "symbolCreateSymbol", self.symbolCreateSymbol)
        vs.SetSavedSetting("importPictures", "symbolFolderSelector", self.symbolFolderSelector)
        vs.SetSavedSetting("importPictures", "symbolFolder", self.symbolFolder)
        vs.SetSavedSetting("importPictures", "importIgnoreErrors", "{}".format(self.importIgnoreErrors))
        vs.SetSavedSetting("importPictures", "importIgnoreExisting", "{}".format(self.importIgnoreExisting))
        vs.SetSavedSetting("importPictures", "importIgnoreUnmodified", "{}".format(self.importIgnoreUnmodified))
