import vs
from vs_constants import *
from _picture_oip import PictureOIP

import pydevd_pycharm
pydevd_pycharm.settrace('localhost', port=12345, stdoutToServer=True, stderrToServer=True, suspend=False)


def execute():
    oip = PictureOIP()
    _, param_name, param_handle, param_rec_handle, wall_handle = vs.GetCustomObjectInfo()

    the_event, the_button = vs.vsoGetEventInfo()

    if the_event == kObjOnInitXProperties:
        oip.create()
    elif the_event == kObjOnWidgetPrep:
        oip.update_parameters_state(param_handle)
    elif the_event == kParametricRecalculate:
        reset_event_handler(param_handle)
        vs.vsoStateClear(param_handle)
    elif the_event == kObjOnAddState:
        _ = vs.vsoStateAddCurrent(param_handle, the_button)


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
