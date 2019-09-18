from typing import Generator, IO
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
                'Driver={{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}};DBQ={};ReadOnly=1;'. \
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

        query_string = 'SELECT * FROM [{}];'.format(self.settings.excelSheetName)

        if self.connected and self.settings.excelSheetName:
            cursor = self.workbook.cursor()
            if cursor:
                criteria_values = []
                for row in cursor.execute(query_string):
                    criteria_values.append(row["{}".format(self.settings.excelCriteriaSelector).lower()])
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

                    picture_name = row["{}".format(self.settings.imageTextureSelector).lower()]
                    if not picture_name:
                        log_message = "UNKNOWN [Error] - Picture name not found"
                        log_file.write(log_message)
                        picture.pictureName = ""
                        return picture
                    else:
                        self.settings.pictureParameters.pictureName = picture_name
                        # Obtain image parameters
                        if self.settings.withImageSelector == "-- Manual":
                            picture.withImage = self.settings.pictureParameters.withImage
                        else:
                            field_value = row["{}".format(self.settings.withImageSelector).lower()]
                            picture.withImage = field_value and field_value != "" and field_value != "False" \
                                and field_value != "No"
                        if picture.withImage:
                            valid, image_width = vs.ValidNumStr(
                                row["{}".format(self.settings.imageWidthSelector).lower()])
                            if valid:
                                picture.imageWidth = str(round(image_width, 3))
                            else:
                                image_message = image_message + "- Invalid Image Width ({})". \
                                    format(picture.imageWidth)
                                valid_picture = False

                            valid, image_height = vs.ValidNumStr(
                                row["{}".format(self.settings.imageHeightSelector).lower()])
                            if valid:
                                picture.imageHeight = str(round(image_height, 3))
                            else:
                                image_message = image_message + "- Invalid Image Height ({})". \
                                    format(picture.imageHeight)
                                valid_picture = False

                            if self.settings.imagePositionSelector == "-- Manual":
                                picture.imagePosition = self.settings.pictureParameters.imagePosition
                                valid = True
                            else:
                                valid, picture.imagePosition = \
                                    vs.ValidNumStr(row["{}".format(self.settings.imagePositionSelector).lower()])
                            if valid:
                                picture.imagePosition = str(round(picture.imagePosition, 3))
                            else:
                                image_message = image_message + "- Invalid Image Position ({})". \
                                    format(picture.imagePosition)
                                valid_picture = False

                        # Obtain frame parameters
                        if self.settings.withFrameSelector == "-- Manual":
                            picture.withFrame = self.settings.pictureParameters.withFrame
                        else:
                            field_value = row["{}".format(self.settings.withFrameSelector).lower()]
                            picture.withFrame = field_value and field_value != "" and field_value != "False" \
                                and field_value != "No"

                        if picture.withFrame == "True":
                            valid, frame_width = vs.ValidNumStr(
                                row["{}".format(self.settings.frameWidthSelector).lower()])
                            if valid:
                                picture.frameWidth = str(round(frame_width, 3))
                            else:
                                frame_message = frame_message + "- Invalid Frame Width ({})". \
                                    format(picture.frameWidth)
                                valid_picture = False

                            valid, frame_height = vs.ValidNumStr(
                                row["{}".format(self.settings.frameHeightSelector).lower()])
                            if valid:
                                picture.frameHeight = str(round(frame_height, 3))
                            else:
                                frame_message = frame_message + "- Invalid Frame Height ({})". \
                                    format(picture.frameHeight)
                                valid_picture = False

                            if self.settings.frameThicknessSelector == "-- Manual":
                                valid, picture.frameThickness = vs.ValidNumStr(
                                    self.settings.pictureParameters.frameThickness)
                                valid = True
                            else:
                                valid, picture.frameThickness = vs.ValidNumStr(
                                    row["{}".format(self.settings.pictureParameters.frameThicknessSelector).lower()])
                            if valid:
                                picture.frameThickness = str(round(picture.frameThickness, 3))
                            else:
                                frame_message = frame_message + "- Invalid Frame Thickness ({})". \
                                    format(picture.frameThickness)
                                valid_picture = False

                            if self.settings.frameDepthSelector == "-- Manual":
                                valid, picture.frameDepth = vs.ValidNumStr(self.settings.pictureParameters.frameDepth)
                                valid = True
                            else:
                                valid, picture.frameDepth = vs.ValidNumStr(
                                    row["{}".format(self.settings.frameDepthSelector).lower()])
                            if valid:
                                picture.frameDepth = str(round(picture.frameDepth, 3))
                            else:
                                frame_message = frame_message + "- Invalid Frame Depth ({})". \
                                    format(picture.frameDepth)
                                valid_picture = False

                            if self.settings.frameClassSelector == "-- Manual":
                                picture.frameClass = self.settings.pictureParameters.frameClass
                            else:
                                picture.frameClass = row["{}".format(self.settings.frameClassSelector).lower()]
                                new_class = vs.GetObject(picture.frameClass)
                                if new_class == 0:
                                    frame_message = frame_message + "- No such Frame Class ({}) ". \
                                        format(picture.frameClass)
                                    valid_picture = False
                                elif vs.ObjectType(new_class) != 94:
                                    frame_message = frame_message + "- Invalid Frame Class ({}) ". \
                                        format(picture.frameClass)
                                    valid_picture = False

                            if self.settings.frameTextureScaleSelector == "-- Manual":
                                picture.frameTextureScale = self.settings.pictureParameters.frameTextureScale
                                valid = True
                            else:
                                valid, picture.frameTextureScale = vs.ValidNumStr(
                                    row["{}".format(self.settings.frameTextureScaleSelector).lower()])
                            if valid:
                                picture.frameTextureScale = str(round(picture.frameTextureScale, 3))
                            else:
                                frame_message = frame_message + "- Invalid Frame Texture Scale ({})". \
                                    format(picture.frameTextureScale)
                                valid_picture = False

                            if self.settings.frameTextureRotationSelector == "-- Manual":
                                picture.frameTextureRotation = self.settings.pictureParameters.frameTextureRotation
                                valid = True
                            else:
                                valid, picture.frameTextureRotation = vs.ValidNumStr(
                                    row["{}".format(self.settings.frameTextureRotationSelector).lower()])
                            if valid:
                                picture.frameTextureRotation = str(round(picture.frameTextureRotation, 3))
                            else:
                                frame_message = frame_message + "- Invalid Frame Texture Rotation ({})". \
                                    format(picture.frameTextureRotation)
                                valid_picture = False

                        # Obtain matboard parameters
                        if self.settings.withMatboardSelector == "-- Manual":
                            picture.withMatboard = self.settings.pictureParameters.withMatboard
                        else:
                            field_value = row["{}".format(self.settings.withMatboardSelector).lower()]
                            picture.withMatboard = field_value and field_value != "" and field_value != "False" \
                                and field_value != "No"

                        if picture.withMatboard == "True":
                            valid, frame_width = vs.ValidNumStr(
                                row["{}".format(self.settings.frameWidthSelector).lower()])
                            if valid:
                                picture.frameWidth = str(round(frame_width, 3))
                            else:
                                frame_message = frame_message + "- Invalid Frame Width ({})".format(picture.frameWidth)
                                valid_picture = False

                            valid, frame_height = vs.ValidNumStr(
                                row["{}".format(self.settings.frameHeightSelector).lower()])
                            if valid:
                                picture.frameHeight = str(round(frame_height, 3))
                            else:
                                frame_message = frame_message + "- Invalid Frame Height ({})". \
                                    format(picture.frameHeight)
                                valid_picture = False

                            if self.settings.matboardPositionSelector == "-- Manual":
                                picture.matboardPosition = self.settings.pictureParameters.matboardPosition
                                valid = True
                            else:
                                valid, picture.matboardPosition = vs.ValidNumStr(
                                    row["{}".format(self.settings.matboardPositionSelector).lower()])
                            if valid:
                                picture.matboardPosition = str(round(picture.matboardPosition, 3))
                            else:
                                matboard_message = matboard_message + "- Invalid Matboard Position ({})". \
                                    format(picture.matboardPosition)
                                valid_picture = False

                            if self.settings.matboardClassSelector == "-- Manual":
                                picture.matboardClass = self.settings.pictureParameters.matboardClass
                            else:
                                picture.matboardClass = row["{}".format(self.settings.matboardClassSelector).lower()]
                                new_class = vs.GetObject(picture.matboardClass)
                                if new_class == 0:
                                    matboard_message = matboard_message + "- No such Matboard Class ({}) ". \
                                        format(picture.matboardClass)
                                    valid_picture = False
                                elif vs.ObjectType(new_class) != 94:
                                    matboard_message = matboard_message + "- Invalid Matboard Class ({})". \
                                        format(picture.matboardClass)
                                    valid_picture = False

                            if self.settings.matboardTextureScaleSelector == "-- Manual":
                                picture.matboardTextureScale = self.settings.pictureParameters.matboardTextureScale
                                valid = True
                            else:
                                valid, picture.matboardTextureScale = vs.ValidNumStr(
                                    row["{}".format(self.settings.matboardTextureScaleSelector).lower()])
                            if valid:
                                picture.matboardTextureScale = str(round(picture.matboardTextureScale, 3))
                            else:
                                matboard_message = matboard_message + "- Invalid Matboard Texture Scale ({})". \
                                    format(picture.matboardTextureScale)
                                valid_picture = False

                            if self.settings.matboardTextureRotatSelector == "-- Manual":
                                picture.matboardTextureRotat = self.settings.pictureParameters.matboardTextureRotat
                                valid = True
                            else:
                                valid, picture.matboardTextureRotat = vs.ValidNumStr(
                                    row["{}".format(self.settings.matboardTextureRotatSelector).lower()])
                            if valid:
                                picture.matboardTextureRotat = str(round(picture.matboardTextureRotat, 3))
                            else:
                                matboard_message = matboard_message + "- Invalid Matboard Texture Rotation ({})". \
                                    format(picture.matboardTextureRotat)
                                valid_picture = False

                        # Obtain glass parameters
                        if self.settings.withGlassSelector == "-- Manual":
                            picture.withGlass = self.settings.pictureParameters.withGlass
                        else:
                            field_value = row["{}".format(self.settings.withGlassSelector).lower()]
                            picture.withGlass = field_value and field_value != "" and field_value != "False" and \
                                field_value != "No"

                        if picture.withGlass == "True":
                            if self.settings.glassPositionSelector == "-- Manual":
                                picture.glassPosition = self.settings.pictureParameters.glassPosition
                                valid = True
                            else:
                                valid, picture.glassPosition = vs.ValidNumStr(
                                    row["{}".format(self.settings.glassPositionSelector).lower()])
                            if valid:
                                picture.glassPosition = str(round(picture.glassPosition, 3))
                            else:
                                glass_message = glass_message + "- Invalid Glass Position ({})". \
                                    format(picture.glassPosition)
                                valid_picture = False

                            if self.settings.glassClassSelector == "-- Manual":
                                picture.glassClass = self.settings.pictureParameters.glassClass
                            else:
                                picture.glassClass = row["{}".format(self.settings.glassClassSelector).lower()]
                                new_class = vs.GetObject(picture.glassClass)
                                if new_class == 0:
                                    glass_message = glass_message + "- No such Glass Class ({}) ". \
                                        format(picture.glassClass)
                                    valid_picture = False
                                elif new_class.type != 94:
                                    glass_message = glass_message + "- Invalid Glass Class ({}) ". \
                                        format(picture.glassClass)
                                    valid_picture = False

                        # Obatian symbol information
                        if self.settings.symbolCreateSymbol == "True":
                            if self.settings.symbolFolderSelector == "-- Manual":
                                picture.symbolFolder = self.settings.symbolFolder
                            else:
                                picture.symbolFolder = row["{}".format(self.settings.symbolFolderSelector).lower()]

                        if not valid_picture:
                            log_message = "{} * [Error]".format(picture_name) + \
                                                        image_message + \
                                                        frame_message + \
                                                        matboard_message + \
                                                        glass_message + "\n"
                            log_file.write(log_message)
                            picture.pictureName = ""

                        yield picture
                cursor.close()
