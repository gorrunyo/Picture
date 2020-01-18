from typing import Generator, IO
import string
import vs
from _picture_settings import PictureParameters
from _import_settings import ImportSettings
import pypyodbc as pyodbc


class ImportDatabase(object):
    """ Picture import workbook Class
    """

    def __init__(self, settings: ImportSettings):
        self.connected = False
        self.workbook = None
        self.settings = settings

    def connect(self) -> bool:
        """ Connects to the excel spreadsheet

        The name of the spreadsheet is specified
        in the settings class member
        
        :returns: True on success. False on failure
        :rtype: bool
        """

        if self.connected:
            self.workbook.close()
            self.connected = False

        if self.settings.excelFileName:
            connection_string = \
                'Driver={{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}};DBQ={};ReadOnly=1;'.\
                format(self.settings.excelFileName)
            try:
                self.workbook = pyodbc.connect(connection_string, autocommit=True)
            except pyodbc.Error as err:
                # vs.SetItemText(importDialog, kWidgetID_excelSheetNameLabel, "Invalid Excel file!")
                vs.AlertCritical(err, "Talk to Carlos")
            else:
                self.connected = True

        return self.connected

    def get_worksheets(self) -> list or None:
        """ Gets the names of all the worksheets in the workbook

        :returns: On success, a list of worksheet names. None on failure
        :rtype: list or None
        """
        if self.connected:
            cursor = self.workbook.cursor()
            if cursor:
                worksheet_names = []
                for table in cursor.tables():
                    worksheet_names.append(table['table_name'])
                cursor.close()
                return worksheet_names
        return None

    def get_columns(self) -> list or None:
        """ Gets the worksheet column names

        The the name of the worksheet is in the `excelSheetName` member
        of `self.settings`

        :returns: On success, a list of sheet column names. None on failure
        :rtype: list or None
        """
        if self.connected and self.settings.excelSheetName:
            cursor = self.workbook.cursor()
            if cursor:
                columns = []
                for row in cursor.columns(self.settings.excelSheetName):
                    columns.append(row['column_name'])
                cursor.close()
                columns.reverse()
                return columns
        return None

    def get_criteria_values(self) -> list or None:
        """ Obtains the criteria values

        Gets the values from the column indicated in `self.settings.excelCriteriaSelector`

        :returns: A list with the critewria values on success and None on failure
        :rtype: list or None
        """

        #        query_string = 'SELECT * FROM [{}];'.format(self.settings.excelSheetName)
        query_string = 'SELECT DISTINCT [{}] FROM [{}];'.format(self.settings.excelCriteriaSelector,
                                                                self.settings.excelSheetName)

        if self.connected and self.settings.excelSheetName:
            cursor = self.workbook.cursor()
            if cursor:
                criteria_values = []
                for row in cursor.execute(query_string):
                    criteria_values.append(row[self.settings.excelCriteriaSelector.lower()])
                cursor.close()
                return criteria_values

        return None

    def get_worksheet_row_count(self) -> int:
        row_count = 0
        if self.connected and self.settings.excelSheetName:
            query_string = 'SELECT * FROM [{}] WHERE [{}] = \'{}\';'.format(self.settings.excelSheetName,
                                                                            self.settings.excelCriteriaSelector,
                                                                            self.settings.excelCriteriaValue)
            cursor = self.workbook.cursor()
            if cursor:
                cursor.execute(query_string)
                rows = cursor.fetchall()
                row_count = len(rows)
                cursor.close()

        return row_count

    def get_worksheet_rows(self, log_file: IO) -> Generator[PictureParameters, None, None]:
        """

        :return:
        """
        picture = PictureParameters()
        if self.connected and self.settings.excelSheetName:
            query_string = 'SELECT * FROM [{}] WHERE [{}] = \'{}\';'.format(self.settings.excelSheetName,
                                                                            self.settings.excelCriteriaSelector,
                                                                            self.settings.excelCriteriaValue)
            cursor = self.workbook.cursor()
            if cursor:
                cursor.execute(query_string)
                rows = cursor.fetchall()
                for row in rows:
                    image_message = ""
                    frame_message = ""
                    matboard_message = ""
                    glass_message = ""
                    valid_picture = True

                    picture_name = row[self.settings.imageTextureSelector.lower()]
                    if not picture_name:
                        log_message = "UNKNOWN [Error] - Picture name not found"
                        log_file.write(log_message)
                        picture.pictureName = ""
                        yield picture
                    else:
                        picture.pictureName = picture_name

                        if picture.pictureName == "Caricature1":
                            stop = True

                        # Obtain image parameters
                        if self.settings.withImageSelector == "-- Manual":
                            picture.withImage = self.settings.pictureParameters.withImage
                        else:
                            cell_value = row[self.settings.withImageSelector.lower()]
                            if cell_value and cell_value != "" and cell_value != "False" and cell_value != "No":
                                picture.withImage = "True"
                            else:
                                picture.withImage = "False"

                        if picture.withImage == "True":
                            cell_value = row[self.settings.imageWidthSelector.lower()]
                            valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                            if valid and value:
                                picture.imageWidth = str(round(value, 3))
                            else:
                                image_message += "- Invalid Image Width ({})".format(cell_value)
                                valid_picture = False

                            cell_value = row[self.settings.imageHeightSelector.lower()]
                            valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                            if valid and value:
                                picture.imageHeight = str(round(value, 3))
                            else:
                                image_message += "- Invalid Image Height ({})".format(cell_value)
                                valid_picture = False

                            if self.settings.imagePositionSelector == "-- Manual":
                                picture.imagePosition = self.settings.pictureParameters.imagePosition
                            else:
                                cell_value = row[self.settings.imagePositionSelector.lower()]
                                valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                                if valid and value:
                                    picture.imagePosition = str(round(value, 3))
                                else:
                                    image_message += "- Invalid Image Position ({})".format(cell_value)
                                    valid_picture = False

                        # Obtain frame parameters
                        if self.settings.withFrameSelector == "-- Manual":
                            picture.withFrame = self.settings.pictureParameters.withFrame
                        else:
                            cell_value = row[self.settings.withFrameSelector.lower()]
                            if cell_value is not None and cell_value != "" and cell_value != "False" and cell_value != "No":
                                picture.withFrame = "True"
                            else:
                                picture.withFrame = "False"

                        if picture.withFrame == "True":
                            cell_value = row[self.settings.frameWidthSelector.lower()]
                            valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                            if valid and value:
                                picture.frameWidth = str(round(value, 3))
                            else:
                                frame_message += "- Invalid Frame Width ({})".format(cell_value)
                                valid_picture = False

                            cell_value = row[self.settings.frameHeightSelector.lower()]
                            valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                            if valid and value:
                                picture.frameHeight = str(round(value, 3))
                            else:
                                frame_message += "- Invalid Frame Height ({})".format(cell_value)
                                valid_picture = False

                            if self.settings.frameThicknessSelector == "-- Manual":
                                picture.frameThickness = self.settings.pictureParameters.frameThickness
                            else:
                                cell_value = row[self.settings.pictureParameters.frameThicknessSelector.lower()]
                                valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                                if valid and value:
                                    picture.frameThickness = str(round(value, 3))
                                else:
                                    frame_message += "- Invalid Frame Thickness ({})".format(cell_value)
                                    valid_picture = False

                            if self.settings.frameDepthSelector == "-- Manual":
                                picture.frameDepth = self.settings.pictureParameters.frameDepth
                            else:
                                cell_value = row[self.settings.frameDepthSelector.lower()]
                                valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                                if valid and value:
                                    picture.frameDepth = str(round(value, 3))
                                else:
                                    frame_message += "- Invalid Frame Depth ({})".format(cell_value)
                                    valid_picture = False

                            if self.settings.frameClassSelector == "-- Manual":
                                picture.frameClass = self.settings.pictureParameters.frameClass
                            else:
                                cell_value = row[self.settings.frameClassSelector.lower()]
                                new_class = vs.GetObject(cell_value)
                                if new_class == 0:
                                    if self.settings.createMissingClasses:
                                        active_class = vs.ActiveClass()
                                        vs.NameClass(cell_value)
                                        vs.NameClass(active_class)
                                        picture.frameClass = cell_value
                                    else:
                                        frame_message += "- No such Frame Class ({})".format(cell_value)
                                        valid_picture = False
                                else:
                                    picture.frameClass = cell_value

                            if self.settings.frameTextureScaleSelector == "-- Manual":
                                picture.frameTextureScale = self.settings.pictureParameters.frameTextureScale
                            else:
                                cell_value = row[self.settings.frameTextureScaleSelector.lower()]
                                valid, picture.frameTextureScale = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                                if valid and value:
                                    picture.frameTextureScale = str(round(value, 3))
                                else:
                                    frame_message += "- Invalid Frame Texture Scale ({})".format(cell_value)
                                    valid_picture = False

                            if self.settings.frameTextureRotationSelector == "-- Manual":
                                picture.frameTextureRotation = self.settings.pictureParameters.frameTextureRotation
                            else:
                                cell_value = row[self.settings.frameTextureRotationSelector.lower()]
                                valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                                if valid and value:
                                    picture.frameTextureRotation = str(round(value, 3))
                                else:
                                    frame_message += "- Invalid Frame Texture Rotation ({})".format(cell_value)
                                    valid_picture = False

                        # Obtain matboard parameters
                        if self.settings.withMatboardSelector == "-- Manual":
                            picture.withMatboard = self.settings.pictureParameters.withMatboard
                        else:
                            cell_value = row[self.settings.withMatboardSelector.lower()]
                            if cell_value and cell_value != "" and cell_value != "False" and cell_value != "No":
                                picture.withMatboard = "True"
                            else:
                                picture.withMatboard = "False"

                        if picture.withMatboard == "True":
                            cell_value = row[self.settings.frameWidthSelector.lower()]
                            valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                            if valid and value:
                                picture.frameWidth = str(round(value, 3))
                            else:
                                frame_message += "- Invalid Frame Width (needed for Matboard) ({})".format(cell_value)
                                valid_picture = False

                            cell_value = row[self.settings.frameHeightSelector.lower()]
                            valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                            if valid and value:
                                picture.frameHeight = str(round(value, 3))
                            else:
                                frame_message += "- Invalid Frame Height (needed for Matboard) ({})".format(cell_value)
                                valid_picture = False

                            cell_value = row[self.settings.windowWidthSelector.lower()]
                            valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                            if valid and value:
                                picture.windowWidth = str(round(value, 3))
                            else:
                                if picture.withImage == "True":
                                    picture.windowWidth = picture.imageWidth
                                    matboard_message += "- Missing window width, using image width instead"
                                else:
                                    matboard_message += "- Invalid Window Width ({})".format(cell_value)
                                    valid_picture = False

                            cell_value = row[self.settings.windowHeightSelector.lower()]
                            valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                            if valid and value:
                                picture.windowHeight = str(round(value, 3))
                            else:
                                if picture.withImage == "True":
                                    picture.windowHeight = picture.imageHeight
                                    matboard_message += "- Missing window height, using image height instead"
                                else:
                                    matboard_message += "- Invalid Window Height ({})".format(cell_value)
                                    valid_picture = False

                            if self.settings.matboardPositionSelector == "-- Manual":
                                picture.matboardPosition = self.settings.pictureParameters.matboardPosition
                            else:
                                cell_value = row[self.settings.matboardPositionSelector.lower()]
                                valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                                if valid and value:
                                    picture.matboardPosition = str(round(value, 3))
                                else:
                                    matboard_message += "- Invalid Matboard Position ({})".format(cell_value)
                                    valid_picture = False

                            if self.settings.matboardClassSelector == "-- Manual":
                                picture.matboardClass = self.settings.pictureParameters.matboardClass
                            else:
                                cell_value = row[self.settings.matboardClassSelector.lower()]
                                new_class = vs.GetObject(cell_value)
                                if new_class == 0:
                                    if self.settings.createMissingClasses:
                                        active_class = vs.ActiveClass()
                                        vs.NameClass(cell_value)
                                        vs.NameClass(active_class)
                                        picture.matboardClass = cell_value
                                    else:
                                        matboard_message += "- No such Matboard Class ({})".format(cell_value)
                                        valid_picture = False
                                else:
                                    picture.matboardClass = cell_value

                            if self.settings.matboardTextureScaleSelector == "-- Manual":
                                picture.matboardTextureScale = self.settings.pictureParameters.matboardTextureScale
                            else:
                                cell_value = row[self.settings.matboardTextureScaleSelector.lower()]
                                valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                                if valid and value:
                                    picture.matboardTextureScale = str(round(value, 3))
                                else:
                                    matboard_message += "- Invalid Matboard Texture Scale ({})".format(cell_value)
                                    valid_picture = False

                            if self.settings.matboardTextureRotatSelector == "-- Manual":
                                picture.matboardTextureRotat = self.settings.pictureParameters.matboardTextureRotat
                            else:
                                cell_value = row[self.settings.matboardTextureRotatSelector.lower()]
                                valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                                if valid and value:
                                    picture.matboardTextureRotat = str(round(value, 3))
                                else:
                                    matboard_message += "- Invalid Matboard Texture Rotation ({})".format(cell_value)
                                    valid_picture = False

                        # Obtain glass parameters
                        if self.settings.withGlassSelector == "-- Manual":
                            picture.withGlass = self.settings.pictureParameters.withGlass
                        else:
                            cell_value = row[self.settings.withGlassSelector.lower()]
                            if cell_value and cell_value != "" and cell_value != "False" and cell_value != "No":
                                picture.withGlass = "True"
                            else:
                                picture.withGlass = "False"

                        if picture.withGlass == "True":
                            if self.settings.glassPositionSelector == "-- Manual":
                                picture.glassPosition = self.settings.pictureParameters.glassPosition
                            else:
                                cell_value = row[self.settings.glassPositionSelector.lower()]
                                valid, value = vs.ValidNumStr(cell_value) if isinstance(cell_value, str) else [True, cell_value]
                                if valid and value:
                                    picture.glassPosition = str(round(value, 3))
                                else:
                                    glass_message += "- Invalid Glass Position ({})".format(cell_value)
                                    valid_picture = False

                            if self.settings.glassClassSelector == "-- Manual":
                                picture.glassClass = self.settings.pictureParameters.glassClass
                            else:
                                cell_value = row[self.settings.glassClassSelector.lower()]
                                new_class = vs.GetObject(picture.glassClass)
                                if new_class == 0:
                                    if self.settings.createMissingClasses:
                                        active_class = vs.ActiveClass()
                                        vs.NameClass(cell_value)
                                        vs.NameClass(active_class)
                                        picture.glassClass = cell_value
                                    else:
                                        glass_message += "- No such Glass Class ({})".format(cell_value)
                                        valid_picture = False
                                else:
                                    picture.glassClass = cell_value

                        # Obtain symbol information
                        if self.settings.symbolCreateSymbol == "True":
                            if self.settings.symbolFolderSelector == "-- Manual":
                                picture.symbolFolder = self.settings.symbolFolder
                            else:
                                folder_name = row[self.settings.symbolFolderSelector.lower()]
                                if folder_name:
                                    picture.symbolFolder = "Picture folder - {}".format(folder_name.translate({ord(c): '_' for c in string.whitespace}).replace("__", "_"))

                        # Obtain Class information
                        if self.settings.classAssignPictureClass == "True":
                            if self.settings.classClassPictureSelector == "-- Manual":
                                picture.pictureClass = self.settings.pictureParameters.pictureClass
                            else:
                                picture.pictureClass = row[self.settings.classClassPictureSelector.lower()]

                        # Obtain Metadata information
                        if self.settings.metaImportMetadata == "True":
                            if picture.withImage == "True":
                                self.settings.pictureRecord.imageSize = "Height: {}, Width: {}".format(picture.imageHeight, picture.imageWidth)
                            if picture.withFrame == "True" or picture.withMatboard == "True":
                                self.settings.pictureRecord.frameSize = "Height: {}, Width: {}".format(picture.frameHeight, picture.frameWidth)
                            if picture.withMatboard == "True":
                                self.settings.pictureRecord.windowSize = "Height: {}, Width: {}".format(picture.windowHeight, picture.windowWidth)
                            if self.settings.metaArtworkTitleSelector != "-- Don't Import":
                                self.settings.pictureRecord.artworkTitle = row[self.settings.metaArtworkTitleSelector.lower()]
                            if self.settings.metaAuthorNameSelector != "-- Don't Import":
                                self.settings.pictureRecord.authorName = row[self.settings.metaAuthorNameSelector.lower()]
                            if self.settings.metaArtworkCreationDateSelector != "-- Don't Import":
                                self.settings.pictureRecord.artworkCreationDate = row[self.settings.metaArtworkCreationDateSelector.lower()]
                            if self.settings.metaArtworkMediaSelector != "-- Don't Import":
                                self.settings.pictureRecord.artworkMedia = row[self.settings.metaArtworkMediaSelector.lower()]
                            # if self.settings.metaTypeSelector != "-- Don't Import":
                            #     self.settings.pictureRecord. = row[self.settings.metaTypeSelector.lower()]
                            # if self.settings.metaRoomLocationSelector != "-- Don't Import":
                            #     self.settings.pictureRecord.roomLocation = row[self.settings.metaRoomLocationSelector.lower()]
                            if self.settings.metaArtworkSourceSelector != "-- Don't Import":
                                self.settings.pictureRecord.artworkSource = row[self.settings.metaArtworkSourceSelector.lower()]
                            if self.settings.metaRegistrationNumberSelector != "-- Don't Import":
                                self.settings.pictureRecord.registrationNumber = row[self.settings.metaRegistrationNumberSelector.lower()]
                            if self.settings.metaAuthorBirthCountrySelector != "-- Don't Import":
                                self.settings.pictureRecord.authorBirthCountry = row[self.settings.metaAuthorBirthCountrySelector.lower()]
                            if self.settings.metaAuthorBirthDateSelector != "-- Don't Import":
                                self.settings.pictureRecord.authorBirthDate = row[self.settings.metaAuthorBirthDateSelector.lower()]
                            if self.settings.metaAuthorDeathDateSelector != "-- Don't Import":
                                self.settings.pictureRecord.authorDeathDate = row[self.settings.metaAuthorDeathDateSelector.lower()]
                            if self.settings.metaDesignNotesSelector != "-- Don't Import":
                                self.settings.pictureRecord.designNotes = row[self.settings.metaDesignNotesSelector.lower()]
                            if self.settings.metaExhibitionMediaSelector != "-- Don't Import":
                                self.settings.pictureRecord.exhibitionMedia = row[self.settings.metaExhibitionMediaSelector.lower()]

                        if not valid_picture:
                            log_message = "{} * [Error]".\
                                              format(picture_name) + image_message + frame_message + matboard_message + glass_message + "\n"
                            log_file.write(log_message)
                            picture.pictureName = ""

                    yield picture
                cursor.close()
