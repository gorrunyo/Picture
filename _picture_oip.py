import vs
from vs_constants import *


class PictureOIP:
    def __init__(self):
        self.kWidgetID_PictureName = 7
        self.kWidgetID_NameEditButton = 8
        self.kWidgetID_ImageSeparator = 9
        self.kWidgetID_WithImage = 10
        self.kWidgetID_ImageWidth = 11
        self.kWidgetID_ImageHeight = 12
        self.kWidgetID_ImagePosition = 13
        self.kWidgetID_ImageTexture = 14
        self.kWidgetID_ImageEditButton = 15
        self.kWidgetID_FrameSeparator = 16
        self.kWidgetID_WithFrame = 17
        self.kWidgetID_FrameWidth = 18
        self.kWidgetID_FrameHeight = 19
        self.kWidgetID_FrameThickness = 20
        self.kWidgetID_FrameDepth = 21
        self.kWidgetID_FrameClass = 22
        self.kWidgetID_FrameTextureScale = 23
        self.kWidgetID_FrameTextureRotation = 24
        self.kWidgetID_MatboardSeparator = 25
        self.kWidgetID_WithMatboard = 26
        self.kWidgetID_MatboardPosition = 27
        self.kWidgetID_MatboardClass = 28
        self.kWidgetID_MatboardTextureScale = 29
        self.kWidgetID_MatboardTextureRotation = 30
        self.kWidgetID_GlassSeparator = 21
        self.kWidgetID_WithGlass = 32
        self.kWidgetID_GlassPosition = 33
        self.kWidgetID_GlassClass = 34

    def create(self):
        _ = vs.SetObjPropVS(kObjXPropHasUIOverride, True)
        _ = vs.SetObjPropVS(kObjXHasCustomWidgetVisibilities, True)
        vs.SetPrefInt(varParametricEnableStateEventing, 1)
        _ = vs.SetObjPropVS(kObjXPropAcceptStates, True)
        _ = vs.SetObjPropVS(kObjXPropAcceptStatesInternal, True)

        _ = vs.vsoAddParamWidget(self.kWidgetID_PictureName, 'PictureName', '')
        _ = vs.vsoAddWidget(self.kWidgetID_ImageSeparator, 100, "Image")
        _ = vs.vsoAddParamWidget(self.kWidgetID_WithImage, 'WithImage', '')
        _ = vs.vsoAddParamWidget(self.kWidgetID_ImageWidth, 'ImageWidth', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_ImageWidth, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_ImageHeight, 'ImageHeight', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_ImageHeight, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_ImagePosition, 'ImagePosition', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_ImagePosition, 1)

        _ = vs.vsoAddWidget(self.kWidgetID_FrameSeparator, 100, "Frame")
        _ = vs.vsoAddParamWidget(self.kWidgetID_WithFrame, 'WithFrame', '')
        _ = vs.vsoAddParamWidget(self.kWidgetID_FrameWidth, 'FrameWidth', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_FrameWidth, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_FrameHeight, 'FrameHeight', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_FrameHeight, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_FrameThickness, 'FrameThickness', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_FrameThickness, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_FrameDepth, 'FrameDepth', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_FrameDepth, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_FrameClass, 'FrameClass', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_FrameClass, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_FrameTextureScale, 'FrameTextureScale', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_FrameTextureScale, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_FrameTextureRotation, 'FrameTextureRotation', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_FrameTextureRotation, 1)

        _ = vs.vsoAddWidget(self.kWidgetID_MatboardSeparator, 100, "Matboard")
        _ = vs.vsoAddParamWidget(self.kWidgetID_WithMatboard, 'WithMatboard', '')
        _ = vs.vsoAddParamWidget(self.kWidgetID_MatboardPosition, 'MatboardPosition', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_MatboardPosition, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_MatboardClass, 'MatboardClass', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_MatboardClass, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_MatboardTextureScale, 'MatboardTextureScale', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_MatboardTextureScale, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_MatboardTextureRotation, 'MatboardTextureRotat', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_MatboardTextureRotation, 1)

        _ = vs.vsoAddWidget(self.kWidgetID_GlassSeparator, 100, "Glass")
        _ = vs.vsoAddParamWidget(self.kWidgetID_WithGlass, 'WithGlass', '')
        _ = vs.vsoAddParamWidget(self.kWidgetID_GlassPosition, 'GlassPosition', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_GlassPosition, 1)
        _ = vs.vsoAddParamWidget(self.kWidgetID_GlassClass, 'GlassClass', '')
        vs.vsoWidgetSetIndLvl(self.kWidgetID_GlassClass, 1)

    # this function updates the visibility or enable/disable state of the widgets
    # note: keep this one fast, it is called often
    def update_parameters_state(self, param_handle):
        vs.vsoWidgetSetVisible(self.kWidgetID_PictureName, param_handle == vs.Handle(0))
        # vs.vsoWidgetSetVisible(self.kWidgetID_NameEditButton, param_handle != vs.Handle(0))

        vs.vsoWidgetSetEnable(self.kWidgetID_ImageWidth, vs.PWithImage)
        vs.vsoWidgetSetEnable(self.kWidgetID_ImageHeight, vs.PWithImage)
        vs.vsoWidgetSetEnable(self.kWidgetID_ImagePosition, vs.PWithImage)
        vs.vsoWidgetSetEnable(self.kWidgetID_ImageTexture, vs.PWithImage)

        vs.vsoWidgetSetEnable(self.kWidgetID_MatboardPosition, vs.PWithMatboard)
        vs.vsoWidgetSetEnable(self.kWidgetID_MatboardClass, vs.PWithMatboard)
        vs.vsoWidgetSetEnable(self.kWidgetID_MatboardTextureScale, vs.PWithMatboard)
        vs.vsoWidgetSetEnable(self.kWidgetID_MatboardTextureRotation, vs.PWithMatboard)

        vs.vsoWidgetSetEnable(self.kWidgetID_FrameWidth, vs.PWithFrame)
        vs.vsoWidgetSetEnable(self.kWidgetID_FrameHeight, vs.PWithFrame)
        vs.vsoWidgetSetEnable(self.kWidgetID_FrameThickness, vs.PWithFrame)
        vs.vsoWidgetSetEnable(self.kWidgetID_FrameDepth, vs.PWithFrame)
        vs.vsoWidgetSetEnable(self.kWidgetID_FrameClass, vs.PWithFrame)
        vs.vsoWidgetSetEnable(self.kWidgetID_FrameTextureScale, vs.PWithFrame)
        vs.vsoWidgetSetEnable(self.kWidgetID_FrameTextureRotation, vs.PWithFrame)

        vs.vsoWidgetSetEnable(self.kWidgetID_GlassPosition, vs.PWithGlass)
        vs.vsoWidgetSetEnable(self.kWidgetID_GlassClass, vs.PWithGlass)

        # this is very important! this is how the system knows we've handled this
        vs.vsoSetEventResult(kObjectEventHandled)
