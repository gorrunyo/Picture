import vs
from _picture import PictureParameters
from vs_constants import *

# import pydevd_pycharm
# pydevd_pycharm.settrace('localhost', port=12345, stdoutToServer=True, stderrToServer=True, suspend=False)


class CreatePictureDialog:
    """ Picture creation dialog object """
    def __init__(self, parameters: PictureParameters):
        # Widget IDs
        ################################################################################################################
        self.kWidgetID_nameGroup = 10
        self.kWidgetID_pictureNameLabel = 11
        self.kWidgetID_pictureName = 12
        self.kWidgetID_createSymbol = 13
        self.kWidgetID_imageGroup = 20
        self.kWidgetID_withImage = 21
        self.kWidgetID_imageWidthLabel = 23
        self.kWidgetID_imageWidth = 24
        self.kWidgetID_imageHeightLabel = 25
        self.kWidgetID_imageHeight = 26
        self.kWidgetID_imagePositionLabel = 27
        self.kWidgetID_imagePosition = 28
        self.kWidgetID_imageTextureLabel = 29
        self.kWidgetID_imageTexture = 30
        self.kWidgetID_imageEditButton = 31
        self.kWidgetID_frameGroup = 40
        self.kWidgetID_withFrame = 41
        self.kWidgetID_frameWidthLabel = 43
        self.kWidgetID_frameWidth = 44
        self.kWidgetID_frameHeightLabel = 45
        self.kWidgetID_frameHeight = 46
        self.kWidgetID_frameThicknessLabel = 47
        self.kWidgetID_frameThickness = 48
        self.kWidgetID_frameDepthLabel = 49
        self.kWidgetID_frameDepth = 50
        self.kWidgetID_frameClassLabel = 51
        self.kWidgetID_frameClass = 52
        self.kWidgetID_frameTextureScaleLabel = 53
        self.kWidgetID_frameTextureScale = 54
        self.kWidgetID_frameTextureRotationLabel = 55
        self.kWidgetID_frameTextureRotation = 56
        self.kWidgetID_matboardGroup = 70
        self.kWidgetID_withMatboard = 71
        self.kWidgetID_windowWidthLabel = 97
        self.kWidgetID_windowWidth = 98
        self.kWidgetID_windowHeightLabel = 99
        self.kWidgetID_windowHeight = 100
        self.kWidgetID_matboardPositionLabel = 73
        self.kWidgetID_matboardPosition = 74
        self.kWidgetID_matboardClassLabel = 75
        self.kWidgetID_matboardClass = 76
        self.kWidgetID_matboardTextureScaleLabel = 77
        self.kWidgetID_matboardTextureScale = 78
        self.kWidgetID_matboardTextureRotatLabel = 79
        self.kWidgetID_matboardTextureRotat = 80
        self.kWidgetID_glassGroup = 90
        self.kWidgetID_withGlass = 91
        self.kWidgetID_glassPositionLabel = 93
        self.kWidgetID_glassPosition = 94
        self.kWidgetID_glassClassLabel = 95
        self.kWidgetID_glassClass = 96
        # last id = 100

        # Dialog settings
        ################################################################################################################
        self.parameters = parameters

        # Run the dialog
        ################################################################################################################
        self.dialog = vs.CreateLayout("Create Picture", True, "OK", "Cancel")
        self.dialog_layout()
        self.result = vs.RunLayoutDialog(self.dialog, self.dialog_handler)

    def dialog_layout(self) -> None:
        """ Creates the dialog layout """
        input_field_width = 18
        label_width = 22

        self.dialog = vs.CreateLayout("Picture Options", True, "OK", "Cancel")

        # Name group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_nameGroup, "", False)
        vs.SetFirstLayoutItem(self.dialog, self.kWidgetID_nameGroup)

        # Name
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_pictureNameLabel, "Picture Name: ", -1)
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_nameGroup, self.kWidgetID_pictureNameLabel)
        vs.CreateEditText(self.dialog, self.kWidgetID_pictureName, self.parameters.pictureName, 30)
        vs.SetRightItem(self.dialog, self.kWidgetID_pictureNameLabel, self.kWidgetID_pictureName, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_pictureName, "Type a name here for the Picture")
        # Create Symbol
        # ------------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_createSymbol, "Create Symbol")
        vs.SetBelowItem(self.dialog, self.kWidgetID_pictureNameLabel, self.kWidgetID_createSymbol, 0, 0)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_createSymbol, self.parameters.createSymbol == "True")

        # Picture group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_imageGroup, "Image", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_nameGroup, self.kWidgetID_imageGroup, 0, 0)

        # With Image checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_withImage, "With Image")
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_imageGroup, self.kWidgetID_withImage)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_withImage, self.parameters.withImage == "True")
        # Image Width dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_imageWidthLabel, "Image Width: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_withImage, self.kWidgetID_imageWidthLabel, 1, 0)
        valid, value = vs.ValidNumStr(self.parameters.imageWidth)
        if not valid:
            value = PictureParameters().imageWidth
        vs.CreateEditReal(self.dialog, self.kWidgetID_imageWidth, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_imageWidthLabel, self.kWidgetID_imageWidth, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_imageWidth, "Enter the width of the image here.")
        # Image Height dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_imageHeightLabel, "Image Height: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_imageWidthLabel, self.kWidgetID_imageHeightLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.imageHeight)
        if not valid:
            value = PictureParameters().imageHeight
        vs.CreateEditReal(self.dialog, self.kWidgetID_imageHeight, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_imageHeightLabel, self.kWidgetID_imageHeight, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_imageHeight, "Enter the height of the image here.")
        # Image Position dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_imagePositionLabel, "Image Position: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_imageHeightLabel, self.kWidgetID_imagePositionLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.imagePosition)
        if not valid:
            value = PictureParameters().imagePosition
        vs.CreateEditReal(self.dialog, self.kWidgetID_imagePosition, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_imagePositionLabel, self.kWidgetID_imagePosition, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_imagePosition, "Enter the position (depth) of the image here.")
        # Image Edit button
        # -----------------------------------------------------------------------------------------
        vs.CreatePushButton(self.dialog, self.kWidgetID_imageEditButton, "Edit Image")
        vs.SetBelowItem(self.dialog, self.kWidgetID_imagePositionLabel, self.kWidgetID_imageEditButton, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_imageEditButton, "Allows to edit the picture image")

        # Frame group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_frameGroup, "Frame", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_imageGroup, self.kWidgetID_frameGroup, 0, 0)

        # With Frame checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_withFrame, "With Frame")
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_frameGroup, self.kWidgetID_withFrame)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_withFrame, self.parameters.withFrame == "True")
        # Frame Width dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameWidthLabel, "Frame Width: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_withFrame, self.kWidgetID_frameWidthLabel, 1, 0)
        valid, value = vs.ValidNumStr(self.parameters.frameWidth)
        if not valid:
            value = PictureParameters().frameWidth
        vs.CreateEditReal(self.dialog, self.kWidgetID_frameWidth, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameWidthLabel, self.kWidgetID_frameWidth, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameWidth, "Enter the width of the frame here.")
        # Frame Height dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameHeightLabel, "Frame Height: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameWidthLabel, self.kWidgetID_frameHeightLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.frameHeight)
        if not valid:
            value = PictureParameters().frameHeight
        vs.CreateEditReal(self.dialog, self.kWidgetID_frameHeight, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameHeightLabel, self.kWidgetID_frameHeight, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameHeight, "Enter the height of the frame here.")
        # Frame Thickness dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameThicknessLabel, "Frame Thickness: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameHeightLabel, self.kWidgetID_frameThicknessLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.frameThickness)
        if not valid:
            value = PictureParameters().frameThickness
        vs.CreateEditReal(self.dialog, self.kWidgetID_frameThickness, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameThicknessLabel, self.kWidgetID_frameThickness, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameThickness, "Enter the thickness of the frame here.")
        # Frame Depth dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameDepthLabel, "Frame Depth: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameThicknessLabel, self.kWidgetID_frameDepthLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.frameDepth)
        if not valid:
            value = PictureParameters().frameDepth
        vs.CreateEditReal(self.dialog, self.kWidgetID_frameDepth, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameDepthLabel, self.kWidgetID_frameDepth, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameDepth, "Enter the depth of the frame here.")
        # Frame Class
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameClassLabel, "Frame Class: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameDepthLabel, self.kWidgetID_frameClassLabel, 0, 0)
        vs.CreateClassPullDownMenu(self.dialog, self.kWidgetID_frameClass, input_field_width)
        class_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_frameClass, self.parameters.frameClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_frameClass, class_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameClassLabel, self.kWidgetID_frameClass, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameClass, "Enter the class of the frame here.")
        # Frame Texture scale
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameTextureScaleLabel, "Frame Texture Scale: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameClassLabel, self.kWidgetID_frameTextureScaleLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.frameTextureScale)
        if not valid:
            value = PictureParameters().frameTextureScale
        vs.CreateEditReal(self.dialog, self.kWidgetID_frameTextureScale, 1, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameTextureScaleLabel, self.kWidgetID_frameTextureScale, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameTextureScale, "Enter the scale for the frame texture.")
        # Frame Texture rotation
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_frameTextureRotationLabel,
                            "Frame Texture Rotation: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameTextureScaleLabel,
                        self.kWidgetID_frameTextureRotationLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.frameTextureRotation)
        if not valid:
            value = PictureParameters().frameTextureRotation
        vs.CreateEditReal(self.dialog, self.kWidgetID_frameTextureRotation, 1, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_frameTextureRotationLabel,
                        self.kWidgetID_frameTextureRotation, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_frameTextureRotation, "Enter the scale for the frame texture.")

        # Matboard group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_matboardGroup, "Matboard", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_frameGroup, self.kWidgetID_matboardGroup, 0, 0)

        # With Matboard checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_withMatboard, "With Matboard")
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_matboardGroup, self.kWidgetID_withMatboard)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_withMatboard, self.parameters.withMatboard == "True")

        # Window Width dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_windowWidthLabel, "Window Width: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_withMatboard, self.kWidgetID_windowWidthLabel, 1, 0)
        valid, value = vs.ValidNumStr(self.parameters.windowWidth)
        if not valid:
            value = PictureParameters().windowWidth
        vs.CreateEditReal(self.dialog, self.kWidgetID_windowWidth, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_windowWidthLabel, self.kWidgetID_windowWidth, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_windowWidth, "Enter the width of the matboard window here.")
        # Window Height dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_windowHeightLabel, "Window Height: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_windowWidthLabel, self.kWidgetID_windowHeightLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.windowHeight)
        if not valid:
            value = PictureParameters().windowHeight
        vs.CreateEditReal(self.dialog, self.kWidgetID_windowHeight, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_windowHeightLabel, self.kWidgetID_windowHeight, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_windowHeight, "Enter the height of the matboard window here.")

        # Matboard Position dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_matboardPositionLabel, "Matboard Position: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_windowHeightLabel, self.kWidgetID_matboardPositionLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.matboardPosition)
        if not valid:
            value = PictureParameters().matboardPosition
        vs.CreateEditReal(self.dialog, self.kWidgetID_matboardPosition, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardPositionLabel, self.kWidgetID_matboardPosition, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardPosition, "Enter the position (depth) of the matboard here.")
        # Matboard Class
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_matboardClassLabel, "Matboard Class: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_matboardPositionLabel, self.kWidgetID_matboardClassLabel, 0, 0)
        vs.CreateClassPullDownMenu(self.dialog, self.kWidgetID_matboardClass, input_field_width)
        class_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_matboardClass, self.parameters.matboardClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_matboardClass, class_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardClassLabel, self.kWidgetID_matboardClass, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardClass, "Enter the class of the matboard here.")
        # Matboard Texture scale
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_matboardTextureScaleLabel,
                            "Matboard Texture Scale: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_matboardClassLabel, self.kWidgetID_matboardTextureScaleLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.matboardTextureScale)
        if not valid:
            value = PictureParameters().matboardTextureScale
        vs.CreateEditReal(self.dialog, self.kWidgetID_matboardTextureScale, 1, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardTextureScaleLabel,
                        self.kWidgetID_matboardTextureScale, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardTextureScale, "Enter the scale for the matboard texture.")
        # Matboard Texture rotation
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_matboardTextureRotatLabel, "Matboard Texture Rotation: ",
                            label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_matboardTextureScaleLabel,
                        self.kWidgetID_matboardTextureRotatLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.matboardTextureRotat)
        if not valid:
            value = PictureParameters().matboardTextureRotat
        vs.CreateEditReal(self.dialog, self.kWidgetID_matboardTextureRotat, 1, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_matboardTextureRotatLabel,
                        self.kWidgetID_matboardTextureRotat, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_matboardTextureRotat, "Enter the scale for the matboard texture.")

        # Glass group
        # =========================================================================================
        vs.CreateGroupBox(self.dialog, self.kWidgetID_glassGroup, "Glass", True)
        vs.SetBelowItem(self.dialog, self.kWidgetID_matboardGroup, self.kWidgetID_glassGroup, 0, 0)

        # With Glass checkbox
        # -----------------------------------------------------------------------------------------
        vs.CreateCheckBox(self.dialog, self.kWidgetID_withGlass, "With Glass")
        vs.SetFirstGroupItem(self.dialog, self.kWidgetID_glassGroup, self.kWidgetID_withGlass)
        vs.SetBooleanItem(self.dialog, self.kWidgetID_withGlass, self.parameters.withGlass == "True")
        # Glass Position dimension
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_glassPositionLabel, "Glass Position: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_withGlass, self.kWidgetID_glassPositionLabel, 0, 0)
        valid, value = vs.ValidNumStr(self.parameters.glassPosition)
        if not valid:
            value = PictureParameters().glassPosition
        vs.CreateEditReal(self.dialog, self.kWidgetID_glassPosition, 3, value, input_field_width)
        vs.SetRightItem(self.dialog, self.kWidgetID_glassPositionLabel, self.kWidgetID_glassPosition, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_glassPosition, "Enter the position (depth) of the glass here.")
        # Glass Class
        # -----------------------------------------------------------------------------------------
        vs.CreateStaticText(self.dialog, self.kWidgetID_glassClassLabel, "Glass Class: ", label_width)
        vs.SetBelowItem(self.dialog, self.kWidgetID_glassPositionLabel, self.kWidgetID_glassClassLabel, 0, 0)
        vs.CreateClassPullDownMenu(self.dialog, self.kWidgetID_glassClass, input_field_width)
        class_index = vs.GetPopUpChoiceIndex(self.dialog, self.kWidgetID_glassClass, self.parameters.matboardClass)
        vs.SelectChoice(self.dialog, self.kWidgetID_glassClass, class_index, True)
        vs.SetRightItem(self.dialog, self.kWidgetID_glassClassLabel, self.kWidgetID_glassClass, 0, 0)
        vs.SetHelpText(self.dialog, self.kWidgetID_glassClass, "Enter the class of the glass here.")

    def dialog_handler(self, item: int, data: int) -> int:
        """ Handles dialog control events """
        if item == KDialogInitEvent:
            vs.SetItemText(self.dialog, self.kWidgetID_pictureName, self.parameters.pictureName)

            # Image Widgets
            # ===============================================================================================
            vs.EnableItem(self.dialog, self.kWidgetID_imageWidthLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imageWidth,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imageHeightLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imageHeight,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imagePositionLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imagePosition,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imageEditButton,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))

            # Frame Widgets
            # ===============================================================================================
            vs.EnableItem(self.dialog, self.kWidgetID_frameWidthLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameWidth,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameHeightLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameHeight,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameThicknessLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameThickness,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameDepthLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameDepth,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameClassLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameClass,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScaleLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScale,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotationLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotation,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))

            # Matboard Widgets
            # ===============================================================================================
            vs.EnableItem(self.dialog, self.kWidgetID_matboardPositionLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard))
            vs.EnableItem(self.dialog, self.kWidgetID_matboardPosition,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard))
            vs.EnableItem(self.dialog, self.kWidgetID_matboardClassLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard))
            vs.EnableItem(self.dialog, self.kWidgetID_matboardClass,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard))
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScaleLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard))
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScale,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard))
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotatLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard))
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotat,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard))

            # Glass Widgets
            # ===============================================================================================
            vs.EnableItem(self.dialog, self.kWidgetID_glassPositionLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withGlass))
            vs.EnableItem(self.dialog, self.kWidgetID_glassPosition,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withGlass))
            vs.EnableItem(self.dialog, self.kWidgetID_glassClassLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withGlass))
            vs.EnableItem(self.dialog, self.kWidgetID_glassClass,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withGlass))

        elif item == self.kWidgetID_pictureName:
            self.parameters.pictureName = vs.GetItemText(self.dialog, self.kWidgetID_pictureName)

        elif item == self.kWidgetID_withImage:
            if not data:
                texture = vs.GetObject(self.parameters.imageTexture)
                if texture != 0:
                    vs.DelObject(texture)
                self.parameters.imageTexture = ""
            elif self.parameters.imageTexture == "":
                self.select_image_texture()
                if self.parameters.imageTexture == "":
                    vs.SetBooleanItem(self.dialog, self.kWidgetID_withImage, False)
            vs.EnableItem(self.dialog, self.kWidgetID_imageWidthLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imageWidth,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imageHeightLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imageHeight,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imagePositionLabel,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imagePosition,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
            vs.EnableItem(self.dialog, self.kWidgetID_imageEditButton,
                          vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))

        elif item == self.kWidgetID_imageEditButton:
            self.select_image_texture()

        elif item == self.kWidgetID_withFrame:
            vs.EnableItem(self.dialog, self.kWidgetID_frameWidthLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameWidth, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameHeightLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameHeight, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameThicknessLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameThickness, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameDepthLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameDepth, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameClassLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameClass, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScaleLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureScale, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotationLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_frameTextureRotation, data)

        elif item == self.kWidgetID_withMatboard:
            vs.EnableItem(self.dialog, self.kWidgetID_matboardPositionLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardPosition, data)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardClassLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardClass, data)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScaleLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureScale, data)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotatLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_matboardTextureRotat, data)

        elif item == self.kWidgetID_withGlass:
            vs.EnableItem(self.dialog, self.kWidgetID_glassPositionLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_glassPosition, data)
            vs.EnableItem(self.dialog, self.kWidgetID_glassClassLabel, data)
            vs.EnableItem(self.dialog, self.kWidgetID_glassClass, data)

        elif item == kOK:
            self.parameters.pictureName = vs.GetItemText(self.dialog, self.kWidgetID_pictureName)
            if self.parameters.pictureName == "":
                vs.AlertInform("The picture name cannot be empty. Please add the picture name", "", True)
                item = -1
            elif vs.GetObject(self.parameters.pictureName) != 0:
                vs.AlertInform("The picture name is already in use. Please change the picture name to avoid a conflict",
                               "", True)
                item = -1
            elif vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage) is False and \
                    vs.GetBooleanItem(self.dialog, self.kWidgetID_withFrame) is False and \
                    vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard) is False and \
                    vs.GetBooleanItem(self.dialog, self.kWidgetID_withGlass) is False:
                vs.AlertInform(
                    "This picture contains no elements. \
                    Please select at least one of the components for the picture object",
                    "", True)
                item = -1
            else:
                self.parameters.createSymbol = "{}".format(vs.GetBooleanItem(self.dialog, self.kWidgetID_createSymbol))
                # Image settings
                # ===============================================================================================
                self.parameters.withImage = "{}".format(vs.GetBooleanItem(self.dialog, self.kWidgetID_withImage))
                _, img_width = vs.GetEditReal(self.dialog, self.kWidgetID_imageWidth, 3)
                self.parameters.imageWidth = "{}\"".format(img_width)
                _, img_height = vs.GetEditReal(self.dialog, self.kWidgetID_imageHeight, 3)
                self.parameters.imageHeight = "{}\"".format(img_height)
                _, img_position = vs.GetEditReal(self.dialog, self.kWidgetID_imagePosition, 3)
#                self.parameters.imagePosition = "{}\"".format(img_position)
                self.parameters.imagePosition = str(img_position)
                texture = vs.GetObject(self.parameters.imageTexture)
                if texture != 0:
                    self.parameters.imageTexture = "{} Prop Texture".format(self.parameters.pictureName)
                    vs.SetName(texture, self.parameters.imageTexture)
                # Frame settings
                # ===============================================================================================
                self.parameters.withFrame = "{}".format(vs.GetBooleanItem(self.dialog, self.kWidgetID_withFrame))
                _, frm_width = vs.GetEditReal(self.dialog, self.kWidgetID_frameWidth, 3)
                self.parameters.frameWidth = "{}".format(frm_width)
                _, frm_height = vs.GetEditReal(self.dialog, self.kWidgetID_frameHeight, 3)
                self.parameters.frameHeight = "{}".format(frm_height)
                _, frm_thickness = vs.GetEditReal(self.dialog, self.kWidgetID_frameThickness, 3)
#                self.parameters.frameThickness = "{}".format(frm_thickness)
                self.parameters.frameThickness = str(frm_thickness)
                _, frm_depth = vs.GetEditReal(self.dialog, self.kWidgetID_frameDepth, 3)
#                self.parameters.frameDepth = "{}".format(frm_depth)
                self.parameters.frameDepth = str(frm_depth)
                _, self.parameters.frameClass = vs.GetSelectedChoiceInfo(self.dialog, self.kWidgetID_frameClass, 0)
                _, frm_texture_scale = vs.GetEditReal(self.dialog, self.kWidgetID_frameTextureScale, 1)
#                self.parameters.frameTextureScale = "{}".format(frm_texture_scale)
                self.parameters.frameTextureScale = str(frm_texture_scale)
                _, frm_texture_rotation = vs.GetEditReal(self.dialog, self.kWidgetID_frameTextureRotation, 1)
#                self.parameters.frameTextureRotation = "{}".format(frm_texture_rotation)
                self.parameters.frameTextureRotation = str(frm_texture_rotation)
                # Matboard settings
                # ===============================================================================================
                self.parameters.withMatboard = "{}".format(vs.GetBooleanItem(self.dialog, self.kWidgetID_withMatboard))
                _, matbrd_position = vs.GetEditReal(self.dialog, self.kWidgetID_matboardPosition, 3)
#                self.parameters.matboardPosition = "{}".format(matbrd_position)
                self.parameters.matboardPosition = str(matbrd_position)
                _, self.parameters.matboardClass = vs.GetSelectedChoiceInfo(self.dialog,
                                                                            self.kWidgetID_matboardClass, 0)
                _, matbrd_texture_scale = vs.GetEditReal(self.dialog, self.kWidgetID_matboardTextureScale, 1)
#                self.parameters.matboardTextureScale = "{}".format(matbrd_texture_scale)
                self.parameters.matboardTextureScale = str(matbrd_texture_scale)
                _, matbrd_texture_rotation = vs.GetEditReal(self.dialog, self.kWidgetID_matboardTextureRotat, 1)
#                self.parameters.matboardTextureRotat = "{}".format(matbrd_texture_rotation)
                self.parameters.matboardTextureRotat = str(matbrd_texture_rotation)
                # Glass settings
                # ===============================================================================================
                self.parameters.withGlass = "{}".format(vs.GetBooleanItem(self.dialog, self.kWidgetID_withGlass))
                _, self.parameters.glassPosition = vs.GetEditReal(self.dialog, self.kWidgetID_glassPosition, 3)
#                self.parameters.glassPosition = "{}".format(self.parameters.glassPosition)
                self.parameters.glassPosition = str(self.parameters.glassPosition)
                _, self.parameters.glassClass = vs.GetSelectedChoiceInfo(self.dialog, self.kWidgetID_glassClass, 0)

        elif item == kCancel:
            texture = vs.GetObject(self.parameters.imageTexture)
            if texture != 0:
                vs.DelObject(texture)
            self.parameters.imageTexture = ""

        return item

    def select_image_texture(self) -> None:
        """ Selects the texture image for the picture object s"""

        if self.parameters.imageTexture == "":
            texture = vs.CreateTexture()
        else:
            texture = vs.GetObject(self.parameters.imageTexture)

        vs.EditTexture(texture)
        shader = vs.GetShaderRecord(texture, 1)
        bitmap = vs.GetTextureBitmap(shader)
        if not bitmap:
            vs.DelObject(shader)
            vs.DelObject(texture)
            self.parameters.imageTexture = ""
        else:
            self.parameters.imageTexture = "{} Prop Texture".format(self.parameters.pictureName)
            vs.SetName(texture, self.parameters.imageTexture)
            shader = vs.GetShaderRecord(texture, 1)
            bitmap = vs.GetTextureBitmap(shader)

            vs.SetTexBitRepHoriz(bitmap, False)
            vs.SetTexBitRepVert(bitmap, False)
