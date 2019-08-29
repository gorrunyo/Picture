import vs
from vs_constants import *
import kWidgetID

import pydevd_pycharm
pydevd_pycharm.settrace('localhost', port=12345, stdoutToServer=True, stderrToServer=True, suspend=False)


def execute():
    _, param_name, param_handle, param_rec_handle, wall_handle = vs.GetCustomObjectInfo()

    the_event, the_button = vs.vsoGetEventInfo()

    if the_event == kObjOnInitXProperties:
        # Enable custom shape pane
        _ = vs.SetObjPropVS(kObjXPropHasUIOverride, True)
        _ = vs.SetObjPropVS(kObjXHasCustomWidgetVisibilities, True)
        vs.SetPrefInt(varParametricEnableStateEventing, 1)
        _ = vs.SetObjPropVS(kObjXPropAcceptStates, True)
        _ = vs.SetObjPropVS(kObjXPropAcceptStatesInternal, True)
        init_oip_layout()

    elif the_event == kObjOnWidgetPrep:
        update_parameters_state(param_handle)

    elif the_event == kObjOnObjectUIButtonHit:
        if the_button == kWidgetID.ImageEditButton:
            on_edit_image_button(param_handle)
        elif the_button == kWidgetID.NameEditButton:
            on_change_picture_name_button(param_handle)

    elif the_event == kParametricRecalculate:
        reset_event_handler(param_handle)
        vs.vsoStateClear(param_handle)

    elif the_event == kObjOnAddState:
        _ = vs.vsoStateAddCurrent(param_handle, the_button)


# this function is executed once and
# it defines the shape pane of the parametric object
#
# The shape pane is composed of widgets
# it is a widget connected to a parameter or it is a button widget
def init_oip_layout():
    # the following line will add all parameters as widgets
    # but we dont want that
    # we would like to set it up ourselves
    # ok = vs.vsoInsertAllParams()

    _ = vs.vsoAddParamWidget(kWidgetID.PictureName, 'PictureName', '')

    _ = vs.vsoAddWidget(kWidgetID.ImageSeparator, 100, "Image")
    _ = vs.vsoAddParamWidget(kWidgetID.WithImage, 'WithImage', '')
    _ = vs.vsoAddParamWidget(kWidgetID.ImageWidth, 'ImageWidth', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.ImageWidth, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.ImageHeight, 'ImageHeight', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.ImageHeight, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.ImagePosition, 'ImagePosition', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.ImagePosition, 1)

    _ = vs.vsoAddWidget(kWidgetID.FrameSeparator, 100, "Frame")
    _ = vs.vsoAddParamWidget(kWidgetID.WithFrame, 'WithFrame', '')
    _ = vs.vsoAddParamWidget(kWidgetID.FrameWidth, 'FrameWidth', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.FrameWidth, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.FrameHeight, 'FrameHeight', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.FrameHeight, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.FrameThickness, 'FrameThickness', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.FrameThickness, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.FrameDepth, 'FrameDepth', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.FrameDepth, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.FrameClass, 'FrameClass', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.FrameClass, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.FrameTextureScale, 'FrameTextureScale', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.FrameTextureScale, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.FrameTextureRotation, 'FrameTextureRotation', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.FrameTextureRotation, 1)

    _ = vs.vsoAddWidget(kWidgetID.MatboardSeparator, 100, "Matboard")
    _ = vs.vsoAddParamWidget(kWidgetID.WithMatboard, 'WithMatboard', '')
    _ = vs.vsoAddParamWidget(kWidgetID.MatboardPosition, 'MatboardPosition', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.MatboardPosition, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.MatboardClass, 'MatboardClass', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.MatboardClass, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.MatboardTextureScale, 'MatboardTextureScale', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.MatboardTextureScale, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.MatboardTextureRotation, 'MatboardTextureRotat', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.MatboardTextureRotation, 1)

    _ = vs.vsoAddWidget(kWidgetID.GlassSeparator, 100, "Glass")
    _ = vs.vsoAddParamWidget(kWidgetID.WithGlass, 'WithGlass', '')
    _ = vs.vsoAddParamWidget(kWidgetID.GlassPosition, 'GlassPosition', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.GlassPosition, 1)
    _ = vs.vsoAddParamWidget(kWidgetID.GlassClass, 'GlassClass', '')
    vs.vsoWidgetSetIndLvl(kWidgetID.GlassClass, 1)


# this function updates the visibility or enable/disable state of the widgets
# note: keep this one fast, it is called often
def update_parameters_state(param_handle):
    vs.vsoWidgetSetVisible(kWidgetID.PictureName, param_handle == vs.Handle(0))
    vs.vsoWidgetSetVisible(kWidgetID.NameEditButton, param_handle != vs.Handle(0))

    vs.vsoWidgetSetEnable(kWidgetID.ImageWidth, vs.PWithImage)
    vs.vsoWidgetSetEnable(kWidgetID.ImageHeight, vs.PWithImage)
    vs.vsoWidgetSetEnable(kWidgetID.ImagePosition, vs.PWithImage)
    #    vs.vsoWidgetSetVisible( picture.kWidgetID_ImageTexture, paramHandle != vs.Handle(0))
    vs.vsoWidgetSetEnable(kWidgetID.ImageTexture, vs.PWithImage)
    vs.vsoWidgetSetEnable(kWidgetID.ImageEditButton, vs.PWithImage)

    vs.vsoWidgetSetEnable(kWidgetID.MatboardPosition, vs.PWithMatboard)
    vs.vsoWidgetSetEnable(kWidgetID.MatboardClass, vs.PWithMatboard)
    vs.vsoWidgetSetEnable(kWidgetID.MatboardTextureScale, vs.PWithMatboard)
    vs.vsoWidgetSetEnable(kWidgetID.MatboardTextureRotation, vs.PWithMatboard)

    vs.vsoWidgetSetEnable(kWidgetID.FrameWidth, vs.PWithFrame)
    vs.vsoWidgetSetEnable(kWidgetID.FrameHeight, vs.PWithFrame)
    vs.vsoWidgetSetEnable(kWidgetID.FrameThickness, vs.PWithFrame)
    vs.vsoWidgetSetEnable(kWidgetID.FrameDepth, vs.PWithFrame)
    vs.vsoWidgetSetEnable(kWidgetID.FrameClass, vs.PWithFrame)
    vs.vsoWidgetSetEnable(kWidgetID.FrameTextureScale, vs.PWithFrame)
    vs.vsoWidgetSetEnable(kWidgetID.FrameTextureRotation, vs.PWithFrame)

    vs.vsoWidgetSetEnable(kWidgetID.GlassPosition, vs.PWithGlass)
    vs.vsoWidgetSetEnable(kWidgetID.GlassClass, vs.PWithGlass)

    # this is very important! this is how the system knows we've handled this
    vs.vsoSetEventResult(kObjectEventHandled)


def on_change_picture_name_button(param_handle):
    new_name = vs.StrDialog("New Picture Name", vs.PPictureName)
    if new_name != "":
        if new_name != vs.PPictureName:
            if vs.GetObject(new_name) != 0:
                vs.AlertInform("An Object with That name already exists", "Please, select a different name", True)
            else:
                vs.PPictureName = new_name
                vs.SetRField(param_handle, "Picture", "PictureName", vs.PPictureName)
                vs.SetName(param_handle, vs.PPictureName)
                texture = vs.GetObject(vs.PImageTexture)
                if texture != 0:
                    vs.PImageTexture = "{} Prop Texture".format(vs.PPictureName)
                    vs.SetRField(param_handle, "Picture", "ImageTexture", vs.PImageTexture)
                    vs.SetName(texture, vs.PImageTexture)
                    vs.ResetObject(param_handle)


def on_edit_image_button(param_handle):
    texture = vs.GetObject(vs.PImageTexture)
    if texture == vs.Handle(0):
        texture = vs.CreateTexture()

    vs.EditTexture(texture)
    shader = vs.GetShaderRecord(texture, 1)
    bitmap = vs.GetTextureBitmap(shader)
    if not bitmap:
        vs.DelObject(shader)
        vs.DelObject(texture)
        vs.PImageTexture = ""
        vs.SetRField(param_handle, "Picture", "ImageTexture", vs.PImageTexture)
        vs.PWithImage = False
        vs.SetRField(param_handle, "Picture", "WithImage", vs.PWithImage)

    else:
        vs.PImageTexture = "{} Prop Texture".format(vs.PPictureName)
        vs.SetRField(param_handle, "Picture", "ImageTexture", vs.PImageTexture)
        vs.SetName(texture, vs.PImageTexture)
        vs.ResetObject(param_handle)
        vs.SetTexBitRepHoriz(bitmap, False)
        vs.SetTexBitRepVert(bitmap, False)


# this function is executed when a parameter changes
# it will define the contents of the parametric object
# everything is created around (0,0) which will appear
# at the insertion point of the parametric object in Vectorworks
def reset_event_handler(rename_handle):
    # Create the Image
    if vs.PWithImage:
        image_texture = vs.GetTextureRefN(rename_handle, 0, 0, True)
        if image_texture:
            image_prop = vs.CreateImageProp(vs.PPictureName, image_texture, vs.PImageHeight, vs.PImageWidth,
                                            False, False, False, False, False)
            if image_prop != 0:
                vs.Move3DObj(image_prop, 0, (vs.PFrameDepth / 2) - vs.PImagePosition, 0)
                existing_texture = vs.GetObject("{} Picture Texture".format(vs.GetName(rename_handle)))
                if existing_texture:
                    set_name(existing_texture, "{} Previous Picture Texture".format(vs.GetName(rename_handle)))
                vs.SetName(
                    vs.GetObject(vs.Index2Name(image_texture)),
                    "{} Picture Texture".format(vs.GetName(rename_handle)))
            else:
                vs.SetRField(rename_handle, "Picture", "WithImage", "False")
                vs.DelObject(vs.GetObject(vs.Index2Name(image_texture)))
                vs.AlertCritical("Error creating Picture object", "Close/Open VectorWorks and retry the operation")

    # Create the Frame
    if vs.PWithFrame:
        vs.BeginPoly3D()

        vs.Add3DPt((-1 * vs.PFrameWidth / 2,
                    -1 * (vs.PFrameDepth / 2),
                    -1 * (vs.PFrameHeight - vs.PImageHeight) / 2))

        vs.Add3DPt((vs.PFrameWidth / 2,
                    -1 * (vs.PFrameDepth / 2),
                    -1 * (vs.PFrameHeight - vs.PImageHeight) / 2))

        vs.Add3DPt((vs.PFrameWidth / 2,
                    -1 * (vs.PFrameDepth / 2),
                    vs.PFrameHeight - ((vs.PFrameHeight - vs.PImageHeight) / 2)))

        vs.Add3DPt((-1 * vs.PFrameWidth / 2,
                    -1 * (vs.PFrameDepth / 2),
                    vs.PFrameHeight - ((vs.PFrameHeight - vs.PImageHeight) / 2)))

        vs.Add3DPt((-1 * vs.PFrameWidth / 2,
                    -1 * (vs.PFrameDepth / 2),
                    -1 * (vs.PFrameHeight - vs.PImageHeight) / 2))

        vs.EndPoly3D()
        extrude_path = vs.LNewObj()
        extrude_path = vs.ConvertToNURBS(extrude_path, False)
        vs.Rect((-1 * vs.PFrameThickness, -1 * vs.PFrameDepth), (0, 0))
        extrude_profile = vs.LNewObj()

        frame = vs.ExtrudeAlongPath(extrude_path, extrude_profile)
        vs.DelObject(extrude_path)
        vs.DelObject(extrude_profile)

        vs.SetClass(frame, vs.PFrameClass)
        vs.SetFPatByClass(frame)
        vs.SetFillColorByClass(frame)
        vs.SetLSByClass(frame)
        vs.SetLWByClass(frame)
        vs.SetMarkerByClass(frame)
        vs.SetOpacityByClass(frame)
        vs.SetPenColorByClass(frame)
        vs.SetTextStyleByClass(frame)
        vs.SetTextureRefN(frame, -1, 0, 0)
        vs.SetTexMapRealN(frame, 3, 0, 3, vs.PFrameTextureScale)
        vs.SetTexMapRealN(frame, 3, 0, 4, vs.Deg2Rad(vs.PFrameTextureRotation))

    # Create the Matboard
    if vs.PWithMatboard:
        vs.BeginPoly3D()
        vs.Add3DPt((-1 * vs.PFrameWidth / 2, (vs.PFrameDepth / 2) - vs.PMatboardPosition,
                    -1 * (vs.PFrameHeight - vs.PImageHeight) / 2))
        vs.Add3DPt((vs.PFrameWidth / 2, (vs.PFrameDepth / 2) - vs.PMatboardPosition,
                    -1 * (vs.PFrameHeight - vs.PImageHeight) / 2))
        vs.Add3DPt((vs.PFrameWidth / 2, (vs.PFrameDepth / 2) - vs.PMatboardPosition,
                    vs.PFrameHeight - ((vs.PFrameHeight - vs.PImageHeight) / 2)))
        vs.Add3DPt((-1 * vs.PFrameWidth / 2, (vs.PFrameDepth / 2) - vs.PMatboardPosition,
                    vs.PFrameHeight - ((vs.PFrameHeight - vs.PImageHeight) / 2)))
        vs.Add3DPt((-1 * vs.PFrameWidth / 2, (vs.PFrameDepth / 2) - vs.PMatboardPosition,
                    -1 * (vs.PFrameHeight - vs.PImageHeight) / 2))
        vs.EndPoly3D()
        matboard = vs.LNewObj()

        vs.SetClass(matboard, vs.PMatboardClass)
        vs.SetFPatByClass(matboard)
        vs.SetFillColorByClass(matboard)
        vs.SetLSByClass(matboard)
        vs.SetLWByClass(matboard)
        vs.SetMarkerByClass(matboard)
        vs.SetOpacityByClass(matboard)
        vs.SetPenColorByClass(matboard)
        vs.SetTextStyleByClass(matboard)
        vs.SetTextureRefN(matboard, -1, 0, 0)
        vs.SetTexMapRealN(matboard, 3, 0, 3, vs.PMatboardTextureScale)
        vs.SetTexMapRealN(matboard, 3, 0, 4, vs.Deg2Rad(vs.PMatboardTextureRotat))

    # Create the Glass
    if vs.PWithGlass:
        vs.BeginPoly3D()
        vs.Add3DPt((-1 * vs.PFrameWidth / 2, (vs.PFrameDepth / 2) - vs.PGlassPosition,
                    -1 * (vs.PFrameHeight - vs.PImageHeight) / 2))
        vs.Add3DPt((vs.PFrameWidth / 2, (vs.PFrameDepth / 2) - vs.PGlassPosition,
                    -1 * (vs.PFrameHeight - vs.PImageHeight) / 2))
        vs.Add3DPt((vs.PFrameWidth / 2, (vs.PFrameDepth / 2) - vs.PGlassPosition,
                    vs.PFrameHeight - ((vs.PFrameHeight - vs.PImageHeight) / 2)))
        vs.Add3DPt((-1 * vs.PFrameWidth / 2, (vs.PFrameDepth / 2) - vs.PGlassPosition,
                    vs.PFrameHeight - ((vs.PFrameHeight - vs.PImageHeight) / 2)))
        vs.Add3DPt((-1 * vs.PFrameWidth / 2, (vs.PFrameDepth / 2) - vs.PGlassPosition,
                    -1 * (vs.PFrameHeight - vs.PImageHeight) / 2))
        vs.EndPoly3D()
        glass = vs.LNewObj()

        vs.SetClass(glass, vs.PGlassClass)
        vs.SetFPatByClass(glass)
        vs.SetFillColorByClass(glass)
        vs.SetLSByClass(glass)
        vs.SetLWByClass(glass)
        vs.SetMarkerByClass(glass)
        vs.SetOpacityByClass(glass)
        vs.SetPenColorByClass(glass)
        vs.SetTextStyleByClass(glass)
        vs.SetTextureRefN(glass, -1, 0, 0)


def set_name(object_handle, name):
    final_name = name
    index = 1
    if vs.GetObject(final_name):
        while vs.GetObject(final_name + " - {}".format(index)):
            index += 1
        vs.SetName(object_handle, final_name + " - {}".format(index))
    else:
        vs.SetName(object_handle, final_name)
