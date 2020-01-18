import vs
from _picture_settings import PictureParameters, PictureRecord


class ImportSettings:
    def __init__(self):
        self.pictureParameters = PictureParameters()
        self.pictureRecord = PictureRecord()
        # self.errorString = ""
        # Picture parameters

        valid, self.pictureParameters.withImage = vs.GetSavedSetting("importPictures", "withImage")
        if not valid or (self.pictureParameters.withImage != "True" and self.pictureParameters.withImage != "False"):
            self.pictureParameters.withImage = PictureParameters().withImage

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "imageWidth")[1])
        self.pictureParameters.imageWidth = str(round(value, 3)) if valid else PictureParameters().imageWidth

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "imageHeight")[1])
        self.pictureParameters.imageHeight = str(round(value, 3)) if valid else PictureParameters().imageHeight

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "imagePosition")[1])
        self.pictureParameters.imagePosition = str(round(value, 3)) if valid else PictureParameters().imagePosition

        valid, self.pictureParameters.imageTexutre = vs.GetSavedSetting("importPictures", "imageTexutre")
        if not valid:
            self.pictureParameters.imageTexutre = PictureParameters().imageTexture

        valid, self.pictureParameters.withFrame = vs.GetSavedSetting("importPictures", "withFrame")
        if not valid or (self.pictureParameters.withFrame != "True" and self.pictureParameters.withFrame != "False"):
            self.pictureParameters.withFrame = PictureParameters().withFrame

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "frameWidth")[1])
        self.pictureParameters.frameWidth = str(round(value, 3)) if valid else PictureParameters().frameWidth

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "frameHeight")[1])
        self.pictureParameters.frameHeight = str(round(value, 3)) if valid else PictureParameters().frameHeight

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "frameThickness")[1])
        self.pictureParameters.frameThickness = str(round(value, 3)) if valid else PictureParameters().frameThickness

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "frameDepth")[1])
        self.pictureParameters.frameDepth = str(round(value, 3)) if valid else PictureParameters().frameDepth

        valid, self.pictureParameters.frameClass = vs.GetSavedSetting("importPictures", "frameClass")
        if not valid:
            self.pictureParameters.frameClass = PictureParameters().frameClass

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "frameTextureScale")[1])
        self.pictureParameters.frameTextureScale = str(round(value, 3)) if valid else PictureParameters().frameTextureScale

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "frameTextureRotation")[1])
        self.pictureParameters.frameTextureRotation = str(round(value, 3)) if valid else PictureParameters().frameTextureRotation

        valid, self.pictureParameters.withMatboard = vs.GetSavedSetting("importPictures", "withMatboard")
        if not valid or (self.pictureParameters.withMatboard != "True" and self.pictureParameters.withMatboard != "False"):
            self.pictureParameters.withMatboard = PictureParameters().withMatboard

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "windowWidth")[1])
        self.pictureParameters.windowWidth = str(round(value, 3)) if valid else PictureParameters().windowWidth

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "windowHeight")[1])
        self.pictureParameters.windowHeight = str(round(value, 3)) if valid else PictureParameters().windowHeight

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "matboardPosition")[1])
        self.pictureParameters.matboardPosition = str(round(value, 3)) if valid else PictureParameters().matboardPosition

        valid, self.pictureParameters.matboardClass = vs.GetSavedSetting("importPictures", "matboardClass")
        if not valid:
            self.pictureParameters.matboardClass = PictureParameters().matboardClass

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "matboardTextureScale")[1])
        self.pictureParameters.matboardTextureScale = str(round(value, 3)) if valid else PictureParameters().matboardTextureScale

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "matboardTextureRotat")[1])
        self.pictureParameters.matboardTextureRotat = str(round(value, 3)) if valid else PictureParameters().matboardTextureRotat

        valid, self.pictureParameters.withGlass = vs.GetSavedSetting("importPictures", "withGlass")
        if not valid or (self.pictureParameters.withGlass != "True" and self.pictureParameters.withGlass != "False"):
            self.pictureParameters.withGlass = PictureParameters().withGlass

        valid, value = vs.ValidNumStr(vs.GetSavedSetting("importPictures", "glassPosition")[1])
        self.pictureParameters.glassPosition = str(round(value, 3)) if valid else PictureParameters().glassPosition

        valid, self.pictureParameters.glassClass = vs.GetSavedSetting("importPictures", "glassClass")
        if not valid:
            self.pictureParameters.glassClass = PictureParameters().glassClass

        valid, self.pictureParameters.pictureClass = vs.GetSavedSetting("importPictures", "pictureClass")
        if not valid:
            self.pictureParameters.pictureClass = PictureParameters().pictureClass

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
        valid, self.windowWidthSelector = vs.GetSavedSetting("importPictures", "windowWidthSelector")
        if not valid:
            self.windowWidthSelector = "-- Select column ..."
        valid, self.windowHeightSelector = vs.GetSavedSetting("importPictures", "windowHeightSelector")
        if not valid:
            self.windowHeightSelector = "-- Select column ..."
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
        valid, self.classAssignPictureClass = vs.GetSavedSetting("importPictures", "classAssignPictureClass")
        if not valid:
            self.classAssignPictureClass = "True"
        valid, self.classClassPictureSelector = vs.GetSavedSetting("importPictures", "classClassPictureSelector")
        if not valid:
            self.classClassPictureSelector = "True"
        valid, self.createMissingClasses = vs.GetSavedSetting("importPictures", "createMissingClasses")
        if not valid:
            self.createMissingClasses = "True"
        valid, self.metaImportMetadata = vs.GetSavedSetting("importPictures", "metaImportMetadata")
        if not valid:
            self.metaImportMetadata = "True"
        valid, self.metaArtworkTitleSelector = vs.GetSavedSetting("importPictures", "metaArtworkTitleSelector")
        if not valid:
            self.metaArtworkTitleSelector = "-- Don't Import"
        valid, self.metaAuthorNameSelector = vs.GetSavedSetting("importPictures", "metaAuthorNameSelector")
        if not valid:
            self.metaAuthorNameSelector = "-- Don't Import"
        valid, self.metaArtworkCreationDateSelector = vs.GetSavedSetting("importPictures", "metaArtworkCreationDateSelector")
        if not valid:
            self.metaArtworkCreationDateSelector = "-- Don't Import"
        valid, self.metaArtworkMediaSelector = vs.GetSavedSetting("importPictures", "metaArtworkMediaSelector")
        if not valid:
            self.metaArtworkMediaSelector = "-- Don't Import"
        valid, self.metaArtworkSourceSelector = vs.GetSavedSetting("importPictures", "metaArtworkSourceSelector")
        if not valid:
            self.metaArtworkSourceSelector = "-- Don't Import"
        valid, self.metaRegistrationNumberSelector = vs.GetSavedSetting("importPictures", "metaRegistrationNumberSelector")
        if not valid:
            self.metaRegistrationNumberSelector = "-- Don't Import"
        valid, self.metaAuthorBirthCountrySelector = vs.GetSavedSetting("importPictures", "metaAuthorBirthCountrySelector")
        if not valid:
            self.metaAuthorBirthCountrySelector = "-- Don't Import"
        valid, self.metaAuthorBirthDateSelector = vs.GetSavedSetting("importPictures", "metaAuthorBirthDateSelector")
        if not valid:
            self.metaAuthorBirthDateSelector = "-- Don't Import"
        valid, self.metaAuthorDeathDateSelector = vs.GetSavedSetting("importPictures", "metaAuthorDeathDateSelector")
        if not valid:
            self.metaAuthorDeathDateSelector = "-- Don't Import"
        valid, self.metaDesignNotesSelector = vs.GetSavedSetting("importPictures", "metaDesignNotesSelector")
        if not valid:
            self.metaDesignNotesSelector = "-- Don't Import"
        valid, self.metaExhibitionMediaSelector = vs.GetSavedSetting("importPictures", "metaExhibitionMediaSelector")
        if not valid:
            self.metaExhibitionMediaSelector = "-- Don't Import"
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

        # Picture parameters
        vs.SetSavedSetting("importPictures", "withImage", self.pictureParameters.withImage)
        vs.SetSavedSetting("importPictures", "imageWidth", str(self.pictureParameters.imageWidth))
        vs.SetSavedSetting("importPictures", "imageHeight", str(self.pictureParameters.imageHeight))
        vs.SetSavedSetting("importPictures", "imagePosition", str(self.pictureParameters.imagePosition))
        vs.SetSavedSetting("importPictures", "imageTexutre", self.pictureParameters.imageTexture)
        vs.SetSavedSetting("importPictures", "withFrame", self.pictureParameters.withFrame)
        vs.SetSavedSetting("importPictures", "frameWidth", str(self.pictureParameters.frameWidth))
        vs.SetSavedSetting("importPictures", "frameHeight", str(self.pictureParameters.frameHeight))
        vs.SetSavedSetting("importPictures", "frameThickness", str(self.pictureParameters.frameThickness))
        vs.SetSavedSetting("importPictures", "frameDepth", str(self.pictureParameters.frameDepth))
        vs.SetSavedSetting("importPictures", "frameClass", self.pictureParameters.frameClass)
        vs.SetSavedSetting("importPictures", "frameTextureScale", str(self.pictureParameters.frameTextureScale))
        vs.SetSavedSetting("importPictures", "frameTextureRotation", str(self.pictureParameters.frameTextureRotation))
        vs.SetSavedSetting("importPictures", "withMatboard", self.pictureParameters.withMatboard)
        vs.SetSavedSetting("importPictures", "windowWidth", str(self.pictureParameters.windowWidth))
        vs.SetSavedSetting("importPictures", "windowHeight", str(self.pictureParameters.windowHeight))
        vs.SetSavedSetting("importPictures", "matboardPosition", str(self.pictureParameters.matboardPosition))
        vs.SetSavedSetting("importPictures", "matboardClass", self.pictureParameters.matboardClass)
        vs.SetSavedSetting("importPictures", "matboardTextureScale", str(self.pictureParameters.matboardTextureScale))
        vs.SetSavedSetting("importPictures", "matboardTextureRotat", str(self.pictureParameters.matboardTextureRotat))
        vs.SetSavedSetting("importPictures", "withGlass", self.pictureParameters.withGlass)
        vs.SetSavedSetting("importPictures", "glassPosition", str(self.pictureParameters.glassPosition))
        vs.SetSavedSetting("importPictures", "glassClass", self.pictureParameters.glassClass)
        vs.SetSavedSetting("importPictures", "pictureClass", self.pictureParameters.pictureClass)

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
        vs.SetSavedSetting("importPictures", "windowWidthSelector", self.windowWidthSelector)
        vs.SetSavedSetting("importPictures", "windowHeightSelector", self.windowHeightSelector)
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
        vs.SetSavedSetting("importPictures", "classAssignPictureClass", "{}".format(self.classAssignPictureClass))
        vs.SetSavedSetting("importPictures", "classClassPictureSelector", "{}".format(self.classClassPictureSelector))
        vs.SetSavedSetting("importPictures", "createMissingClasses", "{}".format(self.createMissingClasses))

        vs.SetSavedSetting("importPictures", "metaImportMetadata", self.metaImportMetadata)
        vs.SetSavedSetting("importPictures", "metaArtworkTitleSelector", "{}".format(self.metaArtworkTitleSelector))
        vs.SetSavedSetting("importPictures", "metaAuthorNameSelector", "{}".format(self.metaAuthorNameSelector))
        vs.SetSavedSetting("importPictures", "metaArtworkCreationDateSelector", "{}".format(self.metaArtworkCreationDateSelector))
        vs.SetSavedSetting("importPictures", "metaArtworkMediaSelector", "{}".format(self.metaArtworkMediaSelector))
        vs.SetSavedSetting("importPictures", "metaArtworkSourceSelector", "{}".format(self.metaArtworkSourceSelector))
        vs.SetSavedSetting("importPictures", "metaRegistrationNumberSelector", "{}".format(self.metaRegistrationNumberSelector))
        vs.SetSavedSetting("importPictures", "metaAuthorBirthCountrySelector", "{}".format(self.metaAuthorBirthCountrySelector))
        vs.SetSavedSetting("importPictures", "metaAuthorBirthDateSelector", "{}".format(self.metaAuthorBirthDateSelector))
        vs.SetSavedSetting("importPictures", "metaAuthorDeathDateSelector", "{}".format(self.metaAuthorDeathDateSelector))
        vs.SetSavedSetting("importPictures", "metaDesignNotesSelector", "{}".format(self.metaDesignNotesSelector))
        vs.SetSavedSetting("importPictures", "metaExhibitionMediaSelector", "{}".format(self.metaExhibitionMediaSelector))

        vs.SetSavedSetting("importPictures", "importIgnoreErrors", "{}".format(self.importIgnoreErrors))
        vs.SetSavedSetting("importPictures", "importIgnoreExisting", "{}".format(self.importIgnoreExisting))
        vs.SetSavedSetting("importPictures", "importIgnoreUnmodified", "{}".format(self.importIgnoreUnmodified))
