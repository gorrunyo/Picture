'''
Created on Dec 7, 2016

@author: Carles
'''
import os
import sys
from time import gmtime, strftime
import vs
import pypyodbc as pyodbc
from vs_constants import *
from _import_settings import ImportSettings
from _import_pictures_dialog import ImportPicturesDialog

import pydevd_pycharm
pydevd_pycharm.settrace('localhost', port=12345, stdoutToServer=True, stderrToServer=True, suspend=False)


# global kOK
# global kCancel
# global KDialogInitEvent
# global KDialogTerminateEvent
# global KDialogTimerEvent
#
# global kWidgetID_excelFileGroup
# global kWidgetID_fileNameLabel
# global kWidgetID_fileName
# global kWidgetID_fileBrowseButton
# global kWidgetID_excelSheetGroup
# global kWidgetID_excelSheetNameLabel
# global kWidgetID_excelSheetName
#
# global kWidgetID_imageGroup
# global kWidgetID_withImageLabel
# global kWidgetID_withImageSelector
# global kWidgetID_withImage
# global kWidgetID_imageFolderNameLabel
# global kWidgetID_imageFolderName
# global kWidgetID_imageFolderBrowseButton
# global kWidgetID_imageTextureLabel
# global kWidgetID_imageTextureSelector
# global kWidgetID_imageTexture
# global kWidgetID_imageWidthLabel
# global kWidgetID_imageWidthSelector
# global kWidgetID_imageHeightLabel
# global kWidgetID_imageHeightSelector
# global kWidgetID_imagePositionLabel
# global kWidgetID_imagePositionSelector
# global kWidgetID_imagePosition
#
# global kWidgetID_frameGroup
# global kWidgetID_withFrameLabel
# global kWidgetID_withFrameSelector
# global kWidgetID_withFrame
# global kWidgetID_frameWidthLabel
# global kWidgetID_frameWidthSelector
# global kWidgetID_frameHeightLabel
# global kWidgetID_frameHeightSelector
# global kWidgetID_frameThicknessLabel
# global kWidgetID_frameThicknessSelector
# global kWidgetID_frameThickness
# global kWidgetID_frameDepthLabel
# global kWidgetID_frameDepthSelector
# global kWidgetID_frameDepth
# global kWidgetID_frameClassLabel
# global kWidgetID_frameClassSelector
# global kWidgetID_frameClass
# global kWidgetID_frameTextureScaleLabel
# global kWidgetID_frameTextureScaleSelector
# global kWidgetID_frameTextureScale
# global kWidgetID_frameTextureRotationLabel
# global kWidgetID_frameTextureRotationSelector
# global kWidgetID_frameTextureRotation
#
# global kWidgetID_matboardGroup
# global kWidgetID_withMatboardLabel
# global kWidgetID_withMatboardSelector
# global kWidgetID_withMatboard
# global kWidgetID_matboardPositionLabel
# global kWidgetID_matboardPositionSelector
# global kWidgetID_matboardPosition
# global kWidgetID_matboardClassLabel
# global kWidgetID_matboardClassSelector
# global kWidgetID_matboardClass
# global kWidgetID_matboardTextureScaleLabel
# global kWidgetID_matboardTextureScaleSelector
# global kWidgetID_matboardTextureScale
# global kWidgetID_matboardTextureRotatLabel
# global kWidgetID_matboardTextureRotatSelector
# global kWidgetID_matboardTextureRotat
#
# global kWidgetID_glassGroup
# global kWidgetID_withGlassLabel
# global kWidgetID_withGlassSelector
# global kWidgetID_withGlass
# global kWidgetID_glassPositionLabel
# global kWidgetID_glassPositionSelector
# global kWidgetID_glassPosition
# global kWidgetID_glassClassLabel
# global kWidgetID_glassClassSelector
# global kWidgetID_glassClass
#
# global kWidgetID_excelCriteriaGroup
# global kWidgetID_excelCriteriaLabel
# global kWidgetID_excelCriteriaSelector
# global kWidgetID_excelCriteriaValue
#
# global kWidgetID_symbolGroup
# global kWidgetID_symbolCreateSymbol
# global kWidgetID_symbolFolderLabel
# global kWidgetID_symbolFolderSelector
# global kWidgetID_symbolFolder
#
# global kWidgetID_importGroup
# global kWidgetID_importIgnoreExisting
# global kWidgetID_importIgnoreErrors
# global kWidgetID_importIgnoreUnmodified
# global kWidgetID_importButton
# global kWidgetID_importNewCount
# global kWidgetID_importUpdatedCount
# global kWidgetID_importDeletedCount
# global kWidgetID_importErrorCount
#
# global excelFileName
# global excelSheetName
# global withImage
# global imageFolderName
# global imageWidth
# global imageHeight
# global imagePosition
# global imageTexure
# global withFrame
# global frameWidth
# global frameHeight
# global frameThickness
# global frameDepth
# global frameClass
# global frameTextureScale
# global frameTextureRotation
# global withMatboard
# global matboardPosition
# global matboardClass
# global matboardTextureScale
# global matboardTextureRotat
# global withGlass
# global glassPosition
# global glassClass
#
# global withImageSelector
# global imageTextureSelector
# global imageWidthSelector
# global imageHeightSelector
# global imagePositionSelector
# global withFrameSelector
# global frameWidthSelector
# global frameHeightSelector
# global frameThicknessSelector
# global frameDepthSelector
# global frameClassSelector
# global frameTextureScaleSelector
# global frameTextureRotationSelector
# global withMatboardSelector
# global matboardPositionSelector
# global matboardClassSelector
# global matboardTextureScaleSelector
# global matboardTextureRotatSelector
# global withGlassSelector
# global glassPositionSelector
# global glassClassSelector
#
# global excelCriteriaSelector
# global excelCriteriaValue
#
# global symbolCreateSymbol
# global symbolFolderSelector
# global symbolFolder
#
# global importIgnoreErrors
# global importIgnoreUnmodified
# global importNewCount
# global importUpdatedCount
# global importDeletedCount
# global importErrorCount
#
#
# excelSheetName = "Select an excel sheet"
# withImage = "True"
# imageWidth = 10.0
# imageHeight = 6.0
# imagePosition = 0.3
# imageTexutre = ""
# withFrame = "True"
# frameWidth = 8.0
# frameHeight = 12.0
# frameThickness = 1.0
# frameDepth = 1.0
# frameClass = "None"
# frameTextureScale = 0.1
# frameTextureRotation = 0.0
# withMatboard = "True"
# matboardPosition = 0.25
# matboardClass = "None"
# matboardTextureScale = 0.1
# matboardTextureRotat = 0.0
# withGlass = "True"
# glassPosition = 0.75
# glassClass = "None"
#
# excelFileName = "Enter excel file name"
# withImageSelector = "-- Manual"
# imageFolderName = "Select a folder"
# imageWidthSelector = "-- Select column ..."
# imageHeightSelector = "-- Select column ..."
# imagePositionSelector = "-- Manual"
# imageTextureSelector = "-- Select column ..."
# withFrameSelector = "-- Manual"
# frameWidthSelector = "-- Select column ..."
# frameHeightSelector = "-- Select column ..."
# frameThicknessSelector = "-- Manual"
# frameDepthSelector = "-- Manual"
# frameClassSelector = "-- Manual"
# frameTextureScaleSelector = "-- Manual"
# frameTextureRotationSelector = "-- Manual"
# withMatboardSelector = "-- Manual"
# matboardPositionSelector = "-- Manual"
# matboardClassSelector = "-- Manual"
# matboardTextureScaleSelector = "-- Manual"
# matboardTextureRotatSelector = "-- Manual"
# withGlassSelector = "-- Manual"
# glassPositionSelector = "-- Manual"
# glassClassSelector = "-- Manual"
# excelCriteriaSelector = "-- Select column ..."
# excelCriteriaValue = "-- Select a value ..."
# symbolCreateSymbol = "False"
# symbolFolderSelector = "-- Manual"
# symbolFolder = "Pictures"
# importIgnoreErrors      = "False"
# importIgnoreExisting = "False"
# importIgnoreUnmodified  = "False"
# importNewCount      = 0
# importUpdatedCount  = 0
# importDeletedCount  = 0
# importErrorCount    = 0
#
#
#
#
# global database
#
# kOK                                     = 1
# kCancel                                 = 2
# KDialogInitEvent                        = 12255
# KDialogTerminateEvent                   = 12256
# KDialogTimerEvent                       = 13028
#
# kWidgetID_excelFileGroup                = 10
# kWidgetID_fileNameLabel                 = 11
# kWidgetID_fileName                      = 12
# kWidgetID_fileBrowseButton              = 13
# kWidgetID_excelSheetGroup               = 14
# kWidgetID_excelSheetNameLabel           = 15
# kWidgetID_excelSheetName                = 16
#
# kWidgetID_imageGroup                    = 20
# kWidgetID_withImageLabel                = 21
# kWidgetID_withImageSelector             = 22
# kWidgetID_withImage                     = 23
# kWidgetID_imageFolderNameLabel          = 24
# kWidgetID_imageFolderName               = 25
# kWidgetID_imageFolderBrowseButton       = 26
# kWidgetID_imageTextureLabel             = 27
# kWidgetID_imageTextureSelector          = 28
# kWidgetID_imageWidthLabel               = 29
# kWidgetID_imageWidthSelector            = 30
# kWidgetID_imageHeightLabel              = 31
# kWidgetID_imageHeightSelector           = 32
# kWidgetID_imagePositionLabel            = 33
# kWidgetID_imagePositionSelector         = 34
# kWidgetID_imagePosition                 = 35
#
# kWidgetID_frameGroup                    = 40
# kWidgetID_withFrameLabel                = 41
# kWidgetID_withFrameSelector             = 42
# kWidgetID_withFrame                     = 43
# kWidgetID_frameWidthLabel               = 44
# kWidgetID_frameWidthSelector            = 45
# kWidgetID_frameHeightLabel              = 46
# kWidgetID_frameHeightSelector           = 47
# kWidgetID_frameThicknessLabel           = 48
# kWidgetID_frameThicknessSelector        = 49
# kWidgetID_frameThickness                = 50
# kWidgetID_frameDepthLabel               = 51
# kWidgetID_frameDepthSelector            = 52
# kWidgetID_frameDepth                    = 53
# kWidgetID_frameClassLabel               = 54
# kWidgetID_frameClassSelector            = 55
# kWidgetID_frameClass                    = 56
# kWidgetID_frameTextureScaleLabel        = 57
# kWidgetID_frameTextureScaleSelector     = 58
# kWidgetID_frameTextureScale             = 59
# kWidgetID_frameTextureRotationLabel     = 60
# kWidgetID_frameTextureRotationSelector  = 61
# kWidgetID_frameTextureRotation          = 62
#
# kWidgetID_matboardGroup                 = 70
# kWidgetID_withMatboardLabel             = 71
# kWidgetID_withMatboardSelector          = 72
# kWidgetID_withMatboard                  = 73
# kWidgetID_matboardPositionLabel         = 74
# kWidgetID_matboardPositionSelector      = 75
# kWidgetID_matboardPosition              = 76
# kWidgetID_matboardClassLabel            = 77
# kWidgetID_matboardClassSelector         = 78
# kWidgetID_matboardClass                 = 79
# kWidgetID_matboardTextureScaleLabel     = 80
# kWidgetID_matboardTextureScaleSelector  = 81
# kWidgetID_matboardTextureScale          = 82
# kWidgetID_matboardTextureRotatLabel     = 83
# kWidgetID_matboardTextureRotatSelector  = 84
# kWidgetID_matboardTextureRotat          = 85
#
# kWidgetID_glassGroup                    = 90
# kWidgetID_withGlassLabel                = 91
# kWidgetID_withGlassSelector             = 92
# kWidgetID_withGlass                     = 93
# kWidgetID_glassPositionLabel            = 94
# kWidgetID_glassPositionSelector         = 95
# kWidgetID_glassPosition                 = 96
# kWidgetID_glassClassLabel               = 97
# kWidgetID_glassClassSelector            = 98
# kWidgetID_glassClass                    = 99
#
# kWidgetID_excelCriteriaGroup            = 100
# kWidgetID_excelCriteriaLabel            = 101
# kWidgetID_excelCriteriaSelector         = 102
# kWidgetID_excelCriteriaValue            = 103
#
# kWidgetID_symbolGroup                   = 200
# kWidgetID_symbolCreateSymbol            = 201
# kWidgetID_symbolFolderLabel             = 202
# kWidgetID_symbolFolderSelector          = 203
# kWidgetID_symbolFolder                  = 204
#
# kWidgetID_importGroup                   = 300
# kWidgetID_importIgnoreErrors            = 301
# kWidgetID_importIgnoreExisting          = 302
# kWidgetID_importIgnoreUnmodified        = 303
# kWidgetID_importButton                  = 304
# kWidgetID_importNewCount                = 305
# kWidgetID_importUpdatedCount            = 306
# kWidgetID_importDeletedCount            = 307
# kWidgetID_importErrorCount              = 308
#
# database = 0
#
# def updatePicture(  directory,
#                     pictureName,
#                     withImage,
#                     imageWidth,
#                     imageHeight,
#                     imagePosition,
#                     withFrame,
#                     frameWidth,
#                     frameHeight,
#                     frameThickness,
#                     frameDepth,
#                     frameClass,
#                     frameTextureScale,
#                     frameTextureRotation,
#                     withMatboard,
#                     matboardPosition,
#                     matboardClass,
#                     matboardTextureScale,
#                     matboardTextureRotat,
#                     withGlass,
#                     glassPosition,
#                     glassClass):
#
#     picture = vs.GetObject(pictureName)
#     if  picture == 0:
#         # Create a new Picture Object
#         vs.BeginSym("{} Picture Symbol".format(pictureName))
#         picture = vs.CreateCustomObject("Picture", 0, 0, 0)
#         vs.SetName(picture, pictureName)
#         vs.EndSym()
#         symbol = vs.GetObject("{} Picture Symbol".format(pictureName))
#
#         vs.SetObjectVariableInt(symbol, 1152, 2) #Thumbnail View - Front
#         vs.SetObjectVariableInt(symbol, 1153, 2) #Thumbnail Render - OpenGL
#
#         texture = vs.GetObject("Valve Prop Texture")
#         if texture != 0:
#             newTexture = vs.CreateDuplicateObject(texture, vs.GetParent(texture))
#             if newTexture != 0:
#                 vs.SetName(newTexture, "{} Prop Texture".format(pictureName))
#
#
#     vs.Record(picture, "Picture")
#     vs.Field(picture, "Picture", "PictureName", pictureName)
#     vs.Field(picture, "Picture", "WithImage", withImage)
#     vs.Field(picture, "Picture", "ImageWidth", imageWidth)
#     vs.Field(picture, "Picture", "ImageHeight", imageHeight)
#     vs.Field(picture, "Picture", "ImagePosition", imagePosition)
#     vs.Field(picture, "Picture", "ImageTexture", "{} Prop Texture".format(pictureName))
#     vs.Field(picture, "Picture", "WithFrame", withFrame)
#     vs.Field(picture, "Picture", "FrameWidth", frameWidth)
#     vs.Field(picture, "Picture", "FrameHeight", frameHeight)
#     vs.Field(picture, "Picture", "FrameThickness", frameThickness)
#     vs.Field(picture, "Picture", "FrameDepth", frameDepth)
#     vs.Field(picture, "Picture", "FrameClass", frameClass)
#     vs.Field(picture, "Picture", "FrameTextureScale", frameTextureScale)
#     vs.Field(picture, "Picture", "FrameTextureRotation", frameTextureRotation)
#     vs.Field(picture, "Picture", "WithMatboard", withMatboard)
#     vs.Field(picture, "Picture", "MatboardPosition", matboardPosition)
#     vs.Field(picture, "Picture", "MatboardClass", matboardClass)
#     vs.Field(picture, "Picture", "MatboardTextureScale", matboardTextureScale)
#     vs.Field(picture, "Picture", "MatboardTextureRotat", matboardTextureRotat)
#     vs.Field(picture, "Picture", "WithGlass", withGlass)
#     vs.Field(picture, "Picture", "GlassPosition", glassPosition)
#     vs.Field(picture, "Picture", "GlassClass", glassClass)
#     vs.ResetObject(picture)
#
# def importPictures():
#     global database
#     global excelSheetName
#     global excelCriteriaSelector
#     global excelCriteriaValue
#     global importIgnoreErrors
#     global importIgnoreExisting
#     global importIgnoreUnmodified
#     global importNewCount
#     global importUpdatedCount
#     global importDeletedCount
#     global importErrorCount
#
#     newPictureName = ""
#     newWithImage = ""
#     newImageWidth = 0.0
#     newImageHeight = 0.0
#     newImagePosition = 0.0
#     newImageTexture = ""
#     newWithFrame = ""
#     newFrameWidth = 0.0
#     newFrameHeight = 0.0
#     newFrameThickness = 0.0
#     newFrameDepth = 0.0
#     newFrameClass = ""
#     newFrameTextureScale = ""
#     newFrameTextureRotation = ""
#     newWithMatboard = ""
#     newMatboardPosition = 0.0
#     newMatboardClass = ""
#     newMatboardTextureScale = 0.0
#     newMatboardTextureRotat = 0.0
#     newWithGlass = ""
#     newGlassPosition = 0.0
#     newGlassClass = ""
#     inner = 0
#     outher = 0
#
#
#     queryString = 'SELECT * FROM [{}] WHERE [{}] = \'{}\';'.format(excelSheetName, excelCriteriaSelector, excelCriteriaValue)
#     documentFileName = vs.GetFPathName()
#     documentFolder = os.path.dirname(documentFileName)
#     logFileName = documentFolder + "/" + "Import_Pictures_" + strftime("%y_%m_%d_%H_%M_%S", gmtime()) + ".log"
#
#     logFile = open(logFileName, "w")
#     logFile.write("Start Picture Import: " + strftime("%d / %m / %y at %H : %M : %S", gmtime()) + "\n")
#     logFile.write("--------------------------------------------------------------------------\n")
#
#     active_class = vs.ActiveClass()
#     vs.NameClass("Pictures")
#
#     cursor = database.cursor()
#     if cursor:
#         cursor.execute(queryString)
#         rows = cursor.fetchall()
#         vs.ProgressDlgOpen("Importing Pictures", True)
#         vs.ProgressDlgSetMeter("Importing " + str(len(rows)) + " Pictures ..." );
#         vs.ProgressDlgStart(100.0, len(rows))
#         importNewCount = 0
#         importUpdatedCount = 0
#         importDeletedCount = 0
#         importErrorCount = 0
#
#         for row in rows:
#             validPicture = True
#             message = ""
#             imageMessage = ""
#             frameMessage = ""
#             matboardMessage = ""
#             glassMessage = ""
#
#             if vs.ProgressDlgHasCancel() == True:
#                 break
#             vs.ProgressDlgYield(1)
#             vs.ProgressDlgSetTopMsg("New Pictures: {}".format(importNewCount))
#             vs.ProgressDlgSetBotMsg("Modified Pictures: {}".format(importUpdatedCount))
#
#
#             newPictureName = row["{}".format(imageTextureSelector).lower()]
#             if newPictureName is None or newPictureName == "":
#                 message = "UNKNOWN [Error] - Picture name not found"
#                 validPicture = False
#             else:
#                 existingPicture = vs.GetObject(newPictureName)
#
#                 if withImageSelector == "-- Manual":
#                     newWithImage = withImage
#                 else:
#                     fieldValue = row["{}".format(withImageSelector).lower()]
#                     if fieldValue is None or fieldValue == "" or fieldValue == "False" or fieldValue == "No":
#                         newWithImage = "False"
#                     else:
#                         newWithImage = "True"
#
#                 if newWithImage == "True":
#                     valid, newImageWidth = vs.ValidNumStr(row["{}".format(imageWidthSelector).lower()])
#                     if valid:
#                         newImageWidth = round(newImageWidth, 3)
#                     else:
#                         imageMessage = imageMessage + "- Invalid Image Width "
#                         validPicture = False
#
#                     valid, newImageHeight = vs.ValidNumStr(row["{}".format(imageHeightSelector).lower()])
#                     if valid:
#                         newImageHeight = round(newImageHeight, 3)
#                     else:
#                         imageMessage = imageMessage + "- Invalid Image Height "
#                         validPicture = False
#
#                     if imagePositionSelector == "-- Manual":
#                         newImagePosition = imagePosition
#                         valid = True
#                     else:
#                         valid, newImagePosition = vs.ValidNumStr(row["{}".format(imagePositionSelector).lower()])
#                     if valid:
#                         newImagePosition = round(newImagePosition, 3)
#                     else:
#                         imageMessage = imageMessage + "- Invalid Image Position "
#                         validPicture = False
#
#     #                newImageTexture = row["{}".format(imageTextureSelector).lower()]
#
#                 if withFrameSelector == "-- Manual":
#                     newWithFrame = withFrame
#                 else:
#                     fieldValue = row["{}".format(withFrameSelector).lower()]
#                     if fieldValue is None or fieldValue == "" or fieldValue == "False" or fieldValue == "No":
#                         newWithFrame = "False"
#                     else:
#                         newWithFrame = "True"
#                 if newWithFrame == "True":
#                     valid, newFrameWidth = vs.ValidNumStr(row["{}".format(frameWidthSelector).lower()])
#                     if valid:
#                         newFrameWidth = round(newFrameWidth, 3)
#                     else:
#                         frameMessage = frameMessage + "- Invalid Frame Width "
#                         validPicture = False
#
#                     valid, newFrameHeight = vs.ValidNumStr(row["{}".format(frameHeightSelector).lower()])
#                     if valid:
#                         newFrameHeight = round(newFrameHeight, 3)
#                     else:
#                         frameMessage = frameMessage + "- Invalid Frame Height "
#                         validPicture = False
#
#                     if frameThicknessSelector == "-- Manual":
#                         valid, newFrameThickness = vs.ValidNumStr(frameThickness)
#                     else:
#                         valid, newFrameThickness = vs.ValidNumStr(row["{}".format(frameThicknessSelector).lower()])
#                     if valid:
#                         newFrameThickness = round(newFrameThickness, 3)
#                     else:
#                         frameMessage = frameMessage + "- Invalid Frame Thickness "
#                         validPicture = False
#
#                     if frameDepthSelector == "-- Manual":
#                         valid, newFrameDepth = vs.ValidNumStr(frameDepth)
#                     else:
#                         valid, newFrameDepth = vs.ValidNumStr(row["{}".format(frameDepthSelector).lower()])
#                     if valid:
#                         newFrameDepth = round(newFrameDepth, 3)
#                     else:
#                         frameMessage = frameMessage + "- Invalid Frame Depth "
#                         validPicture = False
#
#                     if frameClassSelector == "-- Manual":
#                         newFrameClass = frameClass
#                     else:
#                         newFrameClass = row["{}".format(frameClassSelector).lower()]
#                     newClass = vs.GetObject(newFrameClass)
#                     if newClass == 0:
#                         frameMessage = frameMessage + "- Invalid Frame Class "
#                         validPicture = False
#                     elif newClass.type != 94:
#                         frameMessage = frameMessage + "- Invalid Frame Class "
#                         validPicture = False
#
#                     if frameTextureScaleSelector == "-- Manual":
#                         newFrameTextureScale = frameTextureScale
#                         valid = True
#                     else:
#                         valid, newFrameTextureScale = vs.ValidNumStr(row["{}".format(frameTextureScaleSelector).lower()])
#                     if valid:
#                         newFrameTextureScale = round(newFrameTextureScale, 3)
#                     else:
#                         frameMessage = frameMessage + "- Invalid Frame Texture Sccale "
#                         validPicture = False
#
#                     if frameTextureRotationSelector == "-- Manual":
#                         newFrameTextureRotation = frameTextureRotation
#                         valid = True
#                     else:
#                         valid, newFrameTextureRotation = vs.ValidNumStr(row["{}".format(frameTextureRotationSelector).lower()])
#                     if valid:
#                         newFrameTextureRotation = round(newFrameTextureRotation, 3)
#                     else:
#                         frameMessage = frameMessage + "- Invalid Frame Texture Rotation "
#                         validPicture = False
#
#                 if withMatboardSelector == "-- Manual":
#                     newWithMatboard = withMatboard
#                 else:
#                     fieldValue = row["{}".format(withMatboardSelector).lower()]
#                     if fieldValue is None or fieldValue == "" or fieldValue == "False" or fieldValue == "No":
#                         newWithMatboard = "False"
#                     else:
#                         newWithMatboard = "True"
#
#                 if newWithMatboard == "True":
#                     valid, newFrameWidth = vs.ValidNumStr(row["{}".format(frameWidthSelector).lower()])
#                     if valid:
#                         newFrameWidth = round(newFrameWidth, 3)
#                     else:
#                         frameMessage = frameMessage + "- Invalid Frame Width "
#                         validPicture = False
#
#                     valid, newFrameHeight = vs.ValidNumStr(row["{}".format(frameHeightSelector).lower()])
#                     if valid:
#                         newFrameHeight = round(newFrameHeight, 3)
#                     else:
#                         frameMessage = frameMessage + "- Invalid Frame Height "
#                         validPicture = False
#
#                     if matboardPositionSelector == "-- Manual":
#                         newMatboardPosition = matboardPosition
#                         valid = True
#                     else:
#                         valid, newMatboardPosition = vs.ValidNumStr(row["{}".format(matboardPositionSelector).lower()])
#                     if valid:
#                         newMatboardPosition = round(newMatboardPosition, 3)
#                     else:
#                         matboardMessage = matboardMessage + "- Invalid Matboard Position "
#                         validPicture = False
#
#                     if matboardClassSelector == "-- Manual":
#                         newMatboardClass = matboardClass
#                     else:
#                         newMatboardClass = row["{}".format(matboardClassSelector).lower()]
#                     newClass = vs.GetObject(newMatboardClass)
#                     if newClass == 0:
#                         matboardMessage = matboardMessage + "- Invalid Matboard Class "
#                         validPicture = False
#                     elif newClass.type != 94:
#                         matboardMessage = matboardMessage + "- Invalid Matboard Class "
#                         validPicture = False
#
#                     if matboardTextureScaleSelector == "-- Manual":
#                         newMatboardTextureScale = matboardTextureScale
#                         valid = True
#                     else:
#                         valid, newMatboardTextureScale = vs.ValidNumStr(row["{}".format(matboardTextureScaleSelector).lower()])
#                     if valid:
#                         newMatboardTextureScale = round(newMatboardTextureScale, 3)
#                     else:
#                         matboardMessage = matboardMessage + "- Invalid Matboard Texture Scale "
#                         validPicture = False
#
#                     if matboardTextureRotatSelector == "-- Manual":
#                         newMatboardTextureRotat = matboardTextureRotat
#                         valid = True
#                     else:
#                         valid, newMatboardTextureRotat = vs.ValidNumStr(row["{}".format(matboardTextureRotatSelector).lower()])
#                     if valid:
#                         newMatboardTextureRotat = round(newMatboardTextureRotat, 3)
#                     else:
#                         matboardMessage = matboardMessage + "- Invalid Matboard Texture Rotation "
#                         validPicture = False
#
#                 if withGlassSelector == "-- Manual":
#                     newWithGlass = withGlass
#                 else:
#                     fieldValue = row["{}".format(withGlassSelector).lower()]
#                     if fieldValue is None or fieldValue == "" or fieldValue == "False" or fieldValue == "No":
#                         newWithGlass = "False"
#                     else:
#                         newWithGlass = "True"
#
#                 if newWithGlass == "True":
#                     if glassPositionSelector == "-- Manual":
#                         newGlassPosition = glassPosition
#                         valid = True
#                     else:
#                         valid, newGlassPosition = vs.ValidNumStr(row["{}".format(glassPositionSelector).lower()])
#                     if valid:
#                         newGlassPosition = round(newGlassPosition, 3)
#                     else:
#                         glassMessage = glassMessage + "- Invalid Glass Position "
#                         validPicture = False
#
#                     if glassClassSelector == "-- Manual":
#                         newGlassClass = glassClass
#                     else:
#                         newGlassClass = row["{}".format(glassClassSelector).lower()]
#                     newClass = vs.GetObject(newGlassClass)
#                     if newClass == 0:
#                         glassMessage = glassMessage + "- Invalid Glass Class "
#                         validPicture = False
#                     elif newClass.type != 94:
#                         glassMessage = glassMessage + "- Invalid Glass Class "
#                         validPicture = False
#
#                 if validPicture:
#                     if existingPicture != 0:
#                         changed = False
#                         if withImageSelector != "-- Manual" or importIgnoreExisting == "False":
#                             existingWithImage = vs.GetRField(existingPicture, "Picture", "WithImage")
#                             if newWithImage != existingWithImage:
#                                 if newWithImage == "True":
#                                     imageMessage = "- Add immage " + imageMessage
#                                 else:
#                                     imageMessage = "- Removed image "
#                                 vs.SetRField(existingPicture, "Picture", "WithImage", newWithImage)
#                                 changed = True
#
#                         if newWithImage == "True":
#                             valid, existingImageWidth = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "ImageWidth"))
#                             existingImageWidth = round(existingImageWidth, 3)
#                             if newImageWidth != existingImageWidth:
#                                 imageMessage = imageMessage + "- Image With changed "
#                                 vs.SetRField(existingPicture, "Picture", "ImageWidth", newImageWidth)
#                                 changed = True
#
#                             valid, existingImageHeight = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "ImageHeight"))
#                             existingImageHeight = round(existingImageHeight, 3)
#                             if newImageHeight != existingImageHeight:
#                                 imageMessage = imageMessage + "- Image Height changed "
#                                 vs.SetRField(existingPicture, "Picture", "ImageHeight", newImageHeight)
#                                 changed = True
#
#                             if imagePositionSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 valid, existingImagePosition = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "ImagePosition"))
#                                 existingImagePosition = round(existingImagePosition, 3)
#                                 if newImagePosition != existingImagePosition:
#                                     imageMessage = imageMessage + "- Image Position changed "
#                                     vs.SetRField(existingPicture, "Picture", "ImagePosition", newImagePosition)
#                                     changed = True
#
#                         if withFrameSelector != "-- Manual" or importIgnoreExisting == "False":
#                             existingWithFrame = vs.GetRField(existingPicture, "Picture", "WithFrame")
#                             if newWithFrame != existingWithFrame:
#                                 if newWithFrame == "True":
#                                     frameMessage = "Add frame " + frameMessage
#                                 else:
#                                     frameMessage = "Removed frame "
#                                 vs.SetRField(existingPicture, "Picture", "WithFrame", newWithFrame)
#                                 changed = True
#
#                         if newWithFrame == "True":
#                             valid, existingFrameWidth = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "FrameWidth"))
#                             existingFrameWidth = round(existingFrameWidth, 3)
#                             if newFrameWidth != existingFrameWidth:
#                                 frameMessage = frameMessage + "- Frame Width changed "
#                                 vs.SetRField(existingPicture, "Picture", "FrameWidth", newFrameWidth)
#                                 changed = True
#
#                             valid, existingFrameHeight = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "FrameHeight"))
#                             existingFrameHeight = round(existingFrameHeight, 3)
#                             if newFrameHeight != existingFrameHeight:
#                                 frameMessage = frameMessage + "- Frame Height changed "
#                                 vs.SetRField(existingPicture, "Picture", "FrameHeight", newFrameHeight)
#                                 changed = True
#
#                             if frameThicknessSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 valid, existingFrameThickness = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "FrameThickness"))
#                                 existingFrameThickness = round(existingFrameThickness, 3)
#                                 if newFrameThickness != existingFrameThickness:
#                                     frameMessage = frameMessage + "- Frame Thickness changed "
#                                     vs.SetRField(existingPicture, "Picture", "FrameThickness", newFrameThickness)
#                                     changed = True
#
#                             if frameDepthSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 valid, existingFrameDepth = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "FrameDepth"))
#                                 existingFrameDepth = round(existingFrameDepth, 3)
#                                 if newFrameDepth != existingFrameDepth:
#                                     frameMessage = frameMessage + "- Frame Depth changed "
#                                     vs.SetRField(existingPicture, "Picture", "FrameDepth", newFrameDepth)
#                                     changed = True
#
#                             if frameClassSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 existingFrameClass = vs.GetRField(existingPicture, "Picture", "FrameClass")
#                                 if newFrameClass != existingFrameClass:
#                                     frameMessage = frameMessage + "- Frame Class changed "
#                                     vs.SetRField(existingPicture, "Picture", "FrameClass", newFrameClass)
#                                     changed = True
#
#                             if frameTextureScaleSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 valid, existingFrameTextureScale = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "FrameTextureScale"))
#                                 existingFrameTextureScale = round(existingFrameTextureScale, 3)
#                                 if newFrameTextureScale != existingFrameTextureScale:
#                                     frameMessage = frameMessage + "- Frame Texture Scale changed "
#                                     vs.SetRField(existingPicture, "Picture", "FrameTextureScale", newFrameTextureScale)
#                                     changed = True
#
#                             if frameTextureRotationSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 valid, existingFrameTextureRotation = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "FrameTextureRotation"))
#                                 existingFrameTextureRotation = round(existingFrameTextureRotation, 3)
#                                 if newFrameTextureRotation != existingFrameTextureRotation:
#                                     frameMessage = frameMessage + "- Frame Texture Rotation changed "
#                                     vs.SetRField(existingPicture, "Picture", "FrameTextureRotation", newFrameTextureRotation)
#                                     changed = True
#
#                         if withMatboardSelector != "-- Manual" or importIgnoreExisting == "False":
#                             existingWithMatboard = vs.GetRField(existingPicture, "Picture", "WithMatboard")
#                             if newWithMatboard != existingWithMatboard:
#                                 if newWithMatboard == "True":
#                                     matboardMessage = "Add matboard " + matboardMessage
#                                 else:
#                                     matboardMessage = "Removed matboard "
#                                 vs.SetRField(existingPicture, "Picture", "WithMatboard", newWithMatboard)
#                                 changed = True
#
#                         if newWithMatboard == "True":
#                             valid, existingFrameWidth = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "FrameWidth"))
#                             existingFrameWidth = round(existingFrameWidth, 3)
#                             if newFrameWidth != existingFrameWidth:
#                                 frameMessage = frameMessage + "- Frame Width changed "
#                                 vs.SetRField(existingPicture, "Picture", "FrameWidth", newFrameWidth)
#                                 changed = True
#
#                             valid, existingFrameHeight = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "FrameHeight"))
#                             existingFrameHeight = round(existingFrameHeight, 3)
#                             if newFrameHeight != existingFrameHeight:
#                                 frameMessage = frameMessage + "- Frame Height changed "
#                                 vs.SetRField(existingPicture, "Picture", "FrameHeight", newFrameHeight)
#                                 changed = True
#
#                             if matboardPositionSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 valid, existingMatboardPosition = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "MatboardPosition"))
#                                 existingMatboardPosition = round(existingMatboardPosition, 3)
#                                 if newMatboardPosition != existingMatboardPosition:
#                                     matboardMessage = matboardMessage + "- Matboard Position changed "
#                                     vs.SetRField(existingPicture, "Picture", "MatboardPosition", newMatboardPosition)
#                                     changed = True
#
#                             if matboardClassSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 existingMatboardClass = vs.GetRField(existingPicture, "Picture", "MatboardClass")
#                                 if newMatboardClass != existingMatboardClass:
#                                     matboardMessage = matboardMessage + "- Matboard Class changed "
#                                     vs.SetRField(existingPicture, "Picture", "MatboardClass", newMatboardClass)
#                                     changed = True
#
#                             if matboardTextureScaleSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 valid, existingMatboardTextureScale = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "MatboardTextureScale"))
#                                 existingMatboardTextureScale = round(existingMatboardTextureScale, 3)
#                                 if newMatboardTextureScale != existingMatboardTextureScale:
#                                     matboardMessage = matboardMessage + "- Matboard Texture Scale changed "
#                                     vs.SetRField(existingPicture, "Picture", "MatboardTextureScale", newMatboardTextureScale)
#                                     changed = True
#
#                             if matboardTextureRotatSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 valid, existingMatboardTextureRotat = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "MatboardTextureRotat"))
#                                 existingMatboardTextureRotat = round(existingMatboardTextureRotat, 3)
#                                 if newMatboardTextureRotat != existingMatboardTextureRotat:
#                                     matboardMessage = matboardMessage + "- Matboard Texture Rotation changed "
#                                     vs.SetRField(existingPicture, "Picture", "MatboardTextureRotat", newMatboardTextureRotat)
#                                     changed = True
#
#
#                         if withGlassSelector != "-- Manual" or importIgnoreExisting == "False":
#                             existingWithGlass = vs.GetRField(existingPicture, "Picture", "WithGlass")
#                             if newWithGlass != existingWithGlass:
#                                 if newWithGlass == "True":
#                                     glassMessage = "Add glass " + imageMessage
#                                 else:
#                                     glassMessage = "Removed glass "
#                                 vs.SetRField(existingPicture, "Picture", "WithGlass", newWithGlass)
#                                 changed = True
#
#                         if newWithGlass == "True":
#                             if glassPositionSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 valid, existingGlassPosition = vs.ValidNumStr(vs.GetRField(existingPicture, "Picture", "GlassPosition"))
#                                 existingGlassPosition = round(existingGlassPosition, 3)
#                                 if newGlassPosition != existingGlassPosition:
#                                     glassMessage = glassMessage + "- Glass Position changed "
#                                     vs.SetRField(existingPicture, "Picture", "GlassPosition", newGlassPosition)
#                                     changed = True
#
#                             if glassClassSelector != "-- Manual" or importIgnoreExisting == "False":
#                                 existingGlassClass = vs.GetRField(existingPicture, "Picture", "GlassClass")
#                                 if newGlassClass != existingGlassClass:
#                                     glassMessage = glassMessage + "- Glass Class changed "
#                                     vs.SetRField(existingPicture, "Picture", "GlassClass", newGlassClass)
#                                     changed = True
#
#                         if changed == True:
#                             vs.ResetObject(existingPicture)
#
#                             message = "{} * [Modified] ".format(newPictureName) + imageMessage + frameMessage + matboardMessage + glassMessage + "\n"
#                             importUpdatedCount += 1
#
#                         else:
#                             if importIgnoreUnmodified != "True":
#                                 message = "{} * [Unmodified] \n".format(newPictureName)
#
#                     # New Picture
#                     elif newWithImage == "True" or newWithFrame == "True" or newWithMatboard == "True" or newWithGlass == "True":
#                         # Create a new Picture Object
#                         if symbolCreateSymbol == "True":
#                             if symbolFolderSelector == "-- Manual":
#                                 folderName = symbolFolder
#                             else:
#                                 folderName = row["{}".format(symbolFolderSelector).lower()]
#                             if folderName != "":
#                                 folder = vs.GetObject(folderName)
#                                 if folder != 0:
#                                     if folder.type != 92:
#                                         folder = 0
#                                 if folder == 0:
#                                     vs.NameObject(folderName)
#                                     vs.BeginFolder()
#                                     vs.EndFolder()
#                                     folder = vs.GetObject(folderName)
#
#                             vs.BeginSym("{} Picture Symbol".format(newPictureName))
#
#                         picture = vs.CreateCustomObjectN("Picture", 0, 0, 0, False)
#                         vs.SetName(picture, newPictureName)
#
#                         if symbolCreateSymbol == "True":
#                             vs.EndSym()
#                             symbol = vs.GetObject("{} Picture Symbol".format(newPictureName))
#                             vs.SetObjectVariableInt(symbol, 1152, 3) #Thumbnail View - Front
#                             vs.SetObjectVariableInt(symbol, 1153, 2) #Thumbnail Render - OpenGL
#                             if folder != 0:
#                                 vs.InsertSymbolInFolder(folder, symbol)
#                                 folder = 0
#                         texture = vs.GetObject("Arroway {}".format(newPictureName.replace('-', ' ').replace('_', ' ')))
#                         if texture == 0:
#                             for outher in range(0, 99):
#                                 for inner in range(1, 99):
#                                     if outher == 0:
#                                         searchName = "Arroway {}".format(newPictureName.replace('-', ' ').replace('_', ' ')) + ' ' + str(inner)
#                                     else:
#                                         searchName = "Arroway {}".format(newPictureName.replace('-', ' ').replace('_', ' ')) + ' ' + str(inner) + ' ' + str(outher)
#                                     texture = vs.GetObject(searchName)
#                                     if texture != 0:
#                                         break
#                                 if texture != 0:
#                                     break
#                         if texture == 0:
#                             texture = vs.CreateTexture()
#                             if texture != 0:
#                                 shader = vs.CreateShaderRecord(texture, 1, 41)
#                                 if shader == 0:
#                                     vs.DelObject(texture)
#                                     message = "{} * [Creation Failed] \n".format(newPictureName)
#                                     texture = 0
#                         if texture != 0:
#                             texture_index = vs.Name2Index(vs.GetName(texture))
#                             vs.SetTextureRefN(picture, texture_index, 0, 0)
#                             vs.SetName(texture, "{} Prop Texture".format(newPictureName))
#                             vs.Record(picture, "Picture")
#                             vs.Field(picture, "Picture", "PictureName", newPictureName)
#                             vs.Field(picture, "Picture", "WithImage", newWithImage)
#                             vs.Field(picture, "Picture", "ImageWidth", str(newImageWidth) + "\"")
#                             vs.Field(picture, "Picture", "ImageHeight", str(newImageHeight) + "\"")
#                             vs.Field(picture, "Picture", "ImagePosition", str(newImagePosition) + "\"")
#                             vs.Field(picture, "Picture", "ImageTexture", "{} Prop Texture".format(newPictureName))
#                             vs.Field(picture, "Picture", "WithFrame", newWithFrame)
#                             vs.Field(picture, "Picture", "FrameWidth", str(newFrameWidth) + "\"")
#                             vs.Field(picture, "Picture", "FrameHeight", str(newFrameHeight) + "\"")
#                             vs.Field(picture, "Picture", "FrameThickness", str(newFrameThickness) + "\"")
#                             vs.Field(picture, "Picture", "FrameDepth", str(newFrameDepth) + "\"")
#                             vs.Field(picture, "Picture", "FrameClass", newFrameClass)
#                             vs.Field(picture, "Picture", "FrameTextureScale", str(newFrameTextureScale))
#                             vs.Field(picture, "Picture", "FrameTextureRotation", str(newFrameTextureRotation))
#                             vs.Field(picture, "Picture", "WithMatboard", newWithMatboard)
#                             vs.Field(picture, "Picture", "MatboardPosition", str(newMatboardPosition) + "\"")
#                             vs.Field(picture, "Picture", "MatboardClass", newMatboardClass)
#                             vs.Field(picture, "Picture", "MatboardTextureScale", str(newMatboardTextureScale))
#                             vs.Field(picture, "Picture", "MatboardTextureRotat", str(newMatboardTextureRotat))
#                             vs.Field(picture, "Picture", "WithGlass", newWithGlass)
#                             vs.Field(picture, "Picture", "GlassPosition", str(newGlassPosition) + "\"")
#                             vs.Field(picture, "Picture", "GlassClass", newGlassClass)
#                             vs.ResetObject(picture)
#                             message = "{} * [New] \n".format(newPictureName)
#                             importNewCount += 1
#
#                 # Invalid
#                 else:
#                     if importIgnoreErrors != "True":
#                         message = "{} * [Error]".format(newPictureName) + imageMessage + frameMessage + matboardMessage + glassMessage + "\n"
#                         importErrorCount += 1
#
#             logFile.write(message)
#         vs.ProgressDlgEnd()
#         vs.ProgressDlgClose()
#     cursor.close
#
#     vs.NameClass(active_class)
#
#     logFile.write("--------------------------------------------------------------------------\n")
#     logFile.write("Total new Pictures: {}\n".format(importNewCount))
#     logFile.write("Total modified Pictures: {}\n".format(importUpdatedCount))
#     logFile.write("Total deleted Pictures: {}\n".format(importDeletedCount))
#     if importIgnoreErrors != "True":
#         logFile.write("Total error Pictures: {}\n".format(importErrorCount))
#     logFile.write("--------------------------------------------------------------------------\n")
#     logFile.close()
#
# def updateCriteriaValue(state):
#     global database
#     global excelSheetName
#     global excelCriteriaSelector
#     global excelCriteriaValue
#
#     queryString = 'SELECT * FROM [{}];'.format(excelSheetName)
#     criteriaValues = set()
#
#     if database and state == True and excelCriteriaSelector != "-- Select column ...":
#         cursor = database.cursor()
#         if cursor:
#             for row in cursor.execute(queryString):
#                 criteriaValues.add(row["{}".format(excelCriteriaSelector).lower()])
#             cursor.close
#             for criteria in criteriaValues:
#                 if criteria:
#                     vs.AddChoice(importDialog, kWidgetID_excelCriteriaValue, criteria, 0)
#             vs.AddChoice(importDialog, kWidgetID_excelCriteriaValue, "Select a value ...", 0)
#             index = vs.GetChoiceIndex(importDialog, kWidgetID_excelCriteriaValue, excelCriteriaValue)
#             if index == -1:
#                 vs.SelectChoice(importDialog, kWidgetID_excelCriteriaValue, 0, True);
#                 excelCriteriaValue = "Select a value ..."
#             else:
#                 vs.SelectChoice(importDialog, kWidgetID_excelCriteriaValue, index, True);
#     else:
#         while vs.GetChoiceCount(importDialog, kWidgetID_excelCriteriaValue):
#             vs.RemoveChoice(importDialog, kWidgetID_excelCriteriaValue, 0)
#
#
# def showParameters(state):
#     global importDialog
#     global excelFileName
#     global database
#     global excelSheetName
#     global withImage
#     global imageFolderName
#     global imageTexure
#     global imageWidth
#     global imageHeight
#     global imagePosition
#     global withFrame
#     global frameWidth
#     global frameHeight
#     global frameThickness
#     global frameDepth
#     global frameClass
#     global frameTextureScale
#     global frameTextureRotation
#     global withMatboard
#     global matboardPosition
#     global matboardClass
#     global matboardTextureScale
#     global matboardTextureRotat
#     global withGlass
#     global glassPosition
#     global glassClass
#
#     global withImageSelector
#     global imageTextureSelector
#     global imageWidthSelector
#     global imageHeightSelector
#     global imagePositionSelector
#     global withFrameSelector
#     global frameWidthSelector
#     global frameHeightSelector
#     global frameThicknessSelector
#     global frameDepthSelector
#     global frameClassSelector
#     global frameTextureScaleSelector
#     global frameTextureRotationSelector
#     global withMatboardSelector
#     global matboardPositionSelector
#     global matboardClassSelector
#     global matboardTextureScaleSelector
#     global matboardTextureRotatSelector
#     global withGlassSelector
#     global glassPositionSelector
#     global glassClassSelector
#     global excelCriteriaSelector
#     global excelCriteriaValue
#     global symbolCreateSymbol
#     global symbolFolderSelector
#     global symbolFolder
#     global importIgnoreErrors
#     global importIgnoreExisting
#     global importIgnoreUnmodified
#
#
#     columns = []
#
#     vs.ShowItem(importDialog, kWidgetID_imageGroup, state)
#     vs.ShowItem(importDialog, kWidgetID_withImageLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_withImageSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_withImage, state)
#     vs.ShowItem(importDialog, kWidgetID_imageFolderNameLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_imageFolderName, state)
#     vs.ShowItem(importDialog, kWidgetID_imageFolderBrowseButton, state)
#     vs.ShowItem(importDialog, kWidgetID_imageWidthLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_imageWidthSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_imageHeightLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_imageHeightSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_imagePositionLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_imagePositionSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_imagePosition, state)
#     vs.ShowItem(importDialog, kWidgetID_imageTextureLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_imageTextureSelector, state)
#
#     vs.ShowItem(importDialog, kWidgetID_frameGroup, state)
#     vs.ShowItem(importDialog, kWidgetID_withFrameLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_withFrameSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_withFrame, state)
#     vs.ShowItem(importDialog, kWidgetID_frameWidthLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_frameWidthSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_frameHeightLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_frameHeightSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_frameThicknessLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_frameThicknessSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_frameThickness, state)
#     vs.ShowItem(importDialog, kWidgetID_frameDepthLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_frameDepthSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_frameDepth, state)
#     vs.ShowItem(importDialog, kWidgetID_frameClassLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_frameClassSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_frameClass, state)
#     vs.ShowItem(importDialog, kWidgetID_frameTextureScaleLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_frameTextureScaleSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_frameTextureScale, state)
#     vs.ShowItem(importDialog, kWidgetID_frameTextureRotationLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_frameTextureRotationSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_frameTextureRotation, state)
#
#     vs.ShowItem(importDialog, kWidgetID_matboardGroup, state)
#     vs.ShowItem(importDialog, kWidgetID_withMatboardLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_withMatboardSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_withMatboard, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardPositionLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardPositionSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardPosition, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardClassLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardClassSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardClass, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardTextureScaleLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardTextureScaleSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardTextureScale, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardTextureRotatLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardTextureRotatSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_matboardTextureRotat, state)
#
#     vs.ShowItem(importDialog, kWidgetID_glassGroup, state)
#     vs.ShowItem(importDialog, kWidgetID_withGlassLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_withGlassSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_withGlass, state)
#     vs.ShowItem(importDialog, kWidgetID_glassPositionLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_glassPositionSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_glassPosition, state)
#     vs.ShowItem(importDialog, kWidgetID_glassClassLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_glassClassSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_glassClass, state)
#
#     vs.ShowItem(importDialog, kWidgetID_excelCriteriaGroup, state)
#     vs.ShowItem(importDialog, kWidgetID_excelCriteriaLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_excelCriteriaSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_excelCriteriaValue, state)
#
#     vs.ShowItem(importDialog, kWidgetID_symbolGroup, state)
#     vs.ShowItem(importDialog, kWidgetID_symbolCreateSymbol, state)
#     vs.ShowItem(importDialog, kWidgetID_symbolFolderLabel, state)
#     vs.ShowItem(importDialog, kWidgetID_symbolFolderSelector, state)
#     vs.ShowItem(importDialog, kWidgetID_symbolFolder, state)
#
#     vs.ShowItem(importDialog, kWidgetID_importGroup, state)
#     vs.ShowItem(importDialog, kWidgetID_importIgnoreErrors, state)
#     vs.ShowItem(importDialog, kWidgetID_importIgnoreExisting, state)
#     vs.ShowItem(importDialog, kWidgetID_importButton, state)
#     vs.ShowItem(importDialog, kWidgetID_importNewCount, state)
#     vs.ShowItem(importDialog, kWidgetID_importUpdatedCount, state)
#     vs.ShowItem(importDialog, kWidgetID_importDeletedCount, state)
#     vs.ShowItem(importDialog, kWidgetID_importErrorCount, state and importIgnoreErrors != "True")
#
#     if state == False:
#         while vs.GetChoiceCount(importDialog, kWidgetID_withImageSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_withImageSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_imageTextureSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_imageTextureSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_imageWidthSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_imageWidthSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_imageHeightSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_imageHeightSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_imagePositionSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_imagePositionSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_withFrameSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_withFrameSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_frameWidthSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_frameWidthSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_frameHeightSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_frameHeightSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_frameThicknessSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_frameThicknessSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_frameDepthSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_frameDepthSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_frameClassSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_frameClassSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_frameTextureScaleSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_frameTextureScaleSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_frameTextureRotationSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_frameTextureRotationSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_withMatboardSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_withMatboardSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_matboardPositionSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_matboardPositionSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_matboardClassSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_matboardClassSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_matboardTextureScaleSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_matboardTextureScaleSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_matboardTextureRotatSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_matboardTextureRotatSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_withGlassSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_withGlassSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_glassPositionSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_glassPositionSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_glassClassSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_glassClassSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_excelCriteriaSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_excelCriteriaSelector, 0)
#         while vs.GetChoiceCount(importDialog, kWidgetID_symbolFolderSelector):
#             vs.RemoveChoice(importDialog, kWidgetID_symbolFolderSelector, 0)
#
#         updateCriteriaValue(False)
#
#     else:
#         cursor = database.cursor()
#         if cursor:
#
#             for row in cursor.columns(excelSheetName):
#                 columns.append(row['column_name'])
#             cursor.close()
#             columns.reverse()
#
#             for column in columns:
#                 vs.AddChoice(importDialog, kWidgetID_withImageSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_imageWidthSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_imageHeightSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_imagePositionSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_imageTextureSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_withFrameSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_frameWidthSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_frameHeightSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_frameThicknessSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_frameDepthSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_frameClassSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_frameTextureScaleSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_frameTextureRotationSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_withMatboardSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_matboardPositionSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_matboardClassSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_matboardTextureScaleSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_matboardTextureRotatSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_withGlassSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_glassPositionSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_glassClassSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_excelCriteriaSelector, column, 0)
#                 vs.AddChoice(importDialog, kWidgetID_symbolFolderSelector, column, 0)
#
#             vs.AddChoice(importDialog, kWidgetID_withImageSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_imageTextureSelector, "-- Select column ...", 0)
#             vs.AddChoice(importDialog, kWidgetID_imageWidthSelector, "-- Select column ...", 0)
#             vs.AddChoice(importDialog, kWidgetID_imageHeightSelector, "-- Select column ...", 0)
#             vs.AddChoice(importDialog, kWidgetID_imagePositionSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_withFrameSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_frameWidthSelector, "-- Select column ...", 0)
#             vs.AddChoice(importDialog, kWidgetID_frameHeightSelector, "-- Select column ...", 0)
#             vs.AddChoice(importDialog, kWidgetID_frameThicknessSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_frameDepthSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_frameClassSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_frameTextureScaleSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_frameTextureRotationSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_withMatboardSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_matboardPositionSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_matboardClassSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_matboardTextureScaleSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_matboardTextureRotatSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_withGlassSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_glassPositionSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_glassClassSelector, "-- Manual", 0)
#             vs.AddChoice(importDialog, kWidgetID_excelCriteriaSelector, "-- Select column ...", 0)
#             vs.AddChoice(importDialog, kWidgetID_symbolFolderSelector, "-- Manual", 0)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_withImageSelector, withImageSelector)
#             vs.SelectChoice(importDialog, kWidgetID_withImageSelector, selectorIndex, True)
#
#             vs.SetBooleanItem(importDialog, kWidgetID_withImage, withImage == "True")
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_imageTextureSelector, imageTextureSelector)
#             vs.SelectChoice(importDialog, kWidgetID_imageTextureSelector, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_imageWidthSelector, imageWidthSelector)
#             vs.SelectChoice(importDialog, kWidgetID_imageWidthSelector, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_imageHeightSelector, imageHeightSelector)
#             vs.SelectChoice(importDialog, kWidgetID_imageHeightSelector, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_imagePositionSelector, imagePositionSelector)
#             vs.SelectChoice(importDialog, kWidgetID_imagePositionSelector, selectorIndex, True)
#
# #            vs.SetItemText(importDialog, kWidgetID_imagePosition, imagePosition)
#             vs.SetEditReal(importDialog, kWidgetID_imagePosition, 3, imagePosition)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_withFrameSelector, withFrameSelector)
#             vs.SelectChoice(importDialog, kWidgetID_withFrameSelector, selectorIndex, True)
#
#             vs.SetBooleanItem(importDialog, kWidgetID_withFrame, withFrame == "True")
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameWidthSelector, frameWidthSelector)
#             vs.SelectChoice(importDialog, kWidgetID_frameWidthSelector, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameHeightSelector, frameHeightSelector)
#             vs.SelectChoice(importDialog, kWidgetID_frameHeightSelector, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameThicknessSelector, frameThicknessSelector)
#             vs.SelectChoice(importDialog, kWidgetID_frameThicknessSelector, selectorIndex, True)
#
# #            vs.SetItemText(importDialog, kWidgetID_frameThickness, frameThickness)
#             vs.SetEditReal(importDialog, kWidgetID_frameThickness, 3, frameThickness)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameDepthSelector, frameDepthSelector)
#             vs.SelectChoice(importDialog, kWidgetID_frameDepthSelector, selectorIndex, True)
#
# #            vs.SetItemText(importDialog, kWidgetID_frameDepth, frameDepth)
#             vs.SetEditReal(importDialog, kWidgetID_frameDepth, 3, frameDepth)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameClassSelector, frameClassSelector)
#             vs.SelectChoice(importDialog, kWidgetID_frameClassSelector, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameClass, frameClass)
#             vs.SelectChoice(importDialog, kWidgetID_frameClass, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameTextureScaleSelector, frameTextureScaleSelector)
#             vs.SelectChoice(importDialog, kWidgetID_frameTextureScaleSelector, selectorIndex, True)
#
# #            vs.SetItemText(importDialog, kWidgetID_frameTextureScale, frameTextureScale)
#             vs.SetEditReal(importDialog, kWidgetID_frameTextureScale, 1, frameTextureScale)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameTextureRotationSelector, frameTextureRotationSelector)
#             vs.SelectChoice(importDialog, kWidgetID_frameTextureRotationSelector, selectorIndex, True)
#
# #            vs.SetItemText(importDialog, kWidgetID_frameTextureRotation, frameTextureRotation)
#             vs.SetEditReal(importDialog, kWidgetID_frameTextureRotation, 1, frameTextureRotation)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_withMatboardSelector, withMatboardSelector)
#             vs.SelectChoice(importDialog, kWidgetID_withMatboardSelector, selectorIndex, True)
#
#             vs.SetBooleanItem(importDialog, kWidgetID_withMatboard, withMatboard == "True")
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_matboardPositionSelector, matboardPositionSelector)
#             vs.SelectChoice(importDialog, kWidgetID_matboardPositionSelector, selectorIndex, True)
#
# #            vs.SetItemText(importDialog, kWidgetID_matboardPosition, matboardPosition)
#             vs.SetEditReal(importDialog, kWidgetID_matboardPosition, 3, matboardPosition)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_matboardClassSelector, matboardClassSelector)
#             vs.SelectChoice(importDialog, kWidgetID_matboardClassSelector, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_matboardClass, matboardClass)
#             vs.SelectChoice(importDialog, kWidgetID_matboardClass, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_matboardTextureScaleSelector, matboardTextureScaleSelector)
#             vs.SelectChoice(importDialog, kWidgetID_matboardTextureScaleSelector, selectorIndex, True)
#
# #            vs.SetItemText(importDialog, kWidgetID_matboardTextureScale, matboardTextureScale)
#             vs.SetEditReal(importDialog, kWidgetID_matboardTextureScale, 1, matboardTextureScale)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_matboardTextureRotatSelector, matboardTextureRotatSelector)
#             vs.SelectChoice(importDialog, kWidgetID_matboardTextureRotatSelector, selectorIndex, True)
#
# #            vs.SetItemText(importDialog, kWidgetID_matboardTextureRotat, matboardTextureRotat)
#             vs.SetEditReal(importDialog, kWidgetID_matboardTextureRotat, 1, matboardTextureRotat)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_withGlassSelector, withGlassSelector)
#             vs.SelectChoice(importDialog, kWidgetID_withGlassSelector, selectorIndex, True)
#
#             vs.SetBooleanItem(importDialog, kWidgetID_withGlass, withGlass == "True")
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_glassPositionSelector, glassPositionSelector)
#             vs.SelectChoice(importDialog, kWidgetID_glassPositionSelector, selectorIndex, True)
#
# #            vs.SetItemText(importDialog, kWidgetID_glassPosition, glassPosition)
#             vs.SetEditReal(importDialog, kWidgetID_glassPosition, 3, glassPosition)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_glassClassSelector, glassClassSelector)
#             vs.SelectChoice(importDialog, kWidgetID_glassClassSelector, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_glassClass, glassClass)
#             vs.SelectChoice(importDialog, kWidgetID_glassClass, selectorIndex, True)
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_excelCriteriaSelector, excelCriteriaSelector)
#             vs.SelectChoice(importDialog, kWidgetID_excelCriteriaSelector, selectorIndex, True)
#
#             vs.SetBooleanItem(importDialog, kWidgetID_symbolCreateSymbol, symbolCreateSymbol == "True")
#
#             selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_symbolFolderSelector, symbolFolderSelector)
#             vs.SelectChoice(importDialog, kWidgetID_symbolFolderSelector, selectorIndex, True)
#
#             updateCriteriaValue(True)
#
#             vs.EnableItem(importDialog, kWidgetID_withImage, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_withImageSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_imagePosition, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_imagePositionSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_withFrame, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_withFrameSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_frameThickness, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_frameThicknessSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_frameDepth, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_frameDepthSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_frameClass, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_frameClassSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_frameTextureScale, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_frameTextureScaleSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_frameTextureRotation, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_frameTextureRotationSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_withMatboard, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_withMatboardSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_matboardPosition, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_matboardPositionSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_matboardClass, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_matboardClassSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_matboardTextureScale, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_matboardTextureScaleSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_matboardTextureRotat, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_matboardTextureRotatSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_withGlass, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_withGlassSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_glassPosition, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_glassPositionSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_glassClass, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_glassClassSelector, 0) == 0)
#             vs.EnableItem(importDialog, kWidgetID_excelCriteriaValue, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_excelCriteriaSelector, 0) != 0)
#             vs.EnableItem(importDialog, kWidgetID_symbolFolder, vs.GetSelectedChoiceIndex(importDialog, kWidgetID_symbolFolderSelector, 0) == 0)
#
#             vs.SetBooleanItem(importDialog, kWidgetID_importIgnoreErrors, importIgnoreErrors == "True")
#             vs.SetBooleanItem(importDialog, kWidgetID_importIgnoreExisting, importIgnoreExisting == "True")
#             vs.SetBooleanItem(importDialog, kWidgetID_importIgnoreUnmodified, importIgnoreUnmodified == "True")
#
#
#
# def importDialogHandler(item, data):
#     global importDialog
#     global excelFileName
#     global database
#     global excelSheetName
#     global withImage
#     global imageFolderName
#     global imageTexure
#     global imageWidth
#     global imageHeight
#     global imagePosition
#     global withFrame
#     global frameWidth
#     global frameHeight
#     global frameThickness
#     global frameDepth
#     global frameClass
#     global frameTextureScale
#     global frameTextureRotation
#     global withMatboard
#     global matboardPosition
#     global matboardClass
#     global matboardTextureScale
#     global matboardTextureRotat
#     global withGlass
#     global glassPosition
#     global glassClass
#
#     global withImageSelector
#     global imageTextureSelector
#     global imageWidthSelector
#     global imageHeightSelector
#     global imagePositionSelector
#     global withFrameSelector
#     global frameWidthSelector
#     global frameHeightSelector
#     global frameThicknessSelector
#     global frameDepthSelector
#     global frameClassSelector
#     global frameTextureScaleSelector
#     global frameTextureRotationSelector
#     global withMatboardSelector
#     global matboardPositionSelector
#     global matboardClassSelector
#     global matboardTextureScaleSelector
#     global matboardTextureRotatSelector
#     global withGlassSelector
#     global glassPositionSelector
#     global glassClassSelector
#     global excelCriteriaSelector
#     global excelCriteriaValue
#
#     global importIgnoreErrors
#     global importIgnoreExisting
#     global importIgnoreUnmodified
#     global symbolCreateSymbol
#     global symbolFolderSelector
#     global symbolFolder
#     global importNewCount
#     global importUpdatedCount
#     global importDeletedCount
#     global importErrorCount
#
#     if item == KDialogInitEvent:
#         vs.SetItemText(importDialog, kWidgetID_fileName, excelFileName)
#
#         vs.SetItemText(importDialog, kWidgetID_imageFolderName, imageFolderName)
#
#         vs.ShowItem(importDialog, kWidgetID_excelSheetNameLabel, False)
#         vs.ShowItem(importDialog, kWidgetID_excelSheetName, False)
#         showParameters(False)
#
#         vs.EnableItem(importDialog, kWidgetID_importButton, False)
#         vs.EnableItem(importDialog, kWidgetID_importNewCount, False)
#         vs.EnableItem(importDialog, kWidgetID_importUpdatedCount, False)
#         vs.EnableItem(importDialog, kWidgetID_importDeletedCount, False)
#
#     elif item == kWidgetID_fileName:
#         excelFileName = vs.GetItemText(importDialog, kWidgetID_fileName)
#
#     elif item == kWidgetID_fileBrowseButton:
#         result, excelFileName = vs.GetFileN("Open Excel file", "", "xlsm")
#         if result:
#             vs.SetItemText(importDialog, kWidgetID_fileName, excelFileName)
#
#     elif item == kWidgetID_excelSheetName:
#         newExcelSheetName = vs.GetChoiceText(importDialog, kWidgetID_excelSheetName, data)
#         if excelSheetName != newExcelSheetName:
#             excelSheetName = newExcelSheetName
#             showParameters(False)
#             if data != 0:
#                 showParameters(True)
#
#     elif item == kWidgetID_withImageSelector:
#         vs.EnableItem(importDialog, kWidgetID_withImage, data == 0)
#         withImageSelector = vs.GetChoiceText(importDialog, kWidgetID_withImageSelector, data)
#     elif item == kWidgetID_withImage:
#         withImage = "{}".format(data == True)
#     elif item == kWidgetID_imageFolderName:
#         imageFolderName = vs.GetItemText(importDialog, kWidgetID_imageFolderName)
#     elif item == kWidgetID_imageFolderBrowseButton:
#         result, imageFolderName = vs.GetFolder("Select the images folder")
#         if result == 0:
#             vs.SetItemText(importDialog, kWidgetID_imageFolderName, imageFolderName)
#     elif item == kWidgetID_imageTextureSelector:
#         imageTextureSelector = vs.GetChoiceText(importDialog, kWidgetID_withImageSelector, data)
#     elif item == kWidgetID_imageWidthSelector:
#         imageWidthSelector = vs.GetChoiceText(importDialog, kWidgetID_imageWidthSelector, data)
#     elif item == kWidgetID_imageHeightSelector:
#         imageHeightSelector = vs.GetChoiceText(importDialog, kWidgetID_imageHeightSelector, data)
#     elif item == kWidgetID_imagePositionSelector:
#         vs.EnableItem(importDialog, kWidgetID_imagePosition, data == 0)
#         imagePositionSelector = vs.GetChoiceText(importDialog, kWidgetID_imagePositionSelector, data)
#     elif item == kWidgetID_imagePosition:
#         _, imagePosition = vs.GetEditReal(importDialog, kWidgetID_imagePosition, 3)
#     elif item == kWidgetID_withFrameSelector:
#         vs.EnableItem(importDialog, kWidgetID_withFrame, data == 0)
#         withFrameSelector = vs.GetChoiceText(importDialog, kWidgetID_withFrameSelector, data)
#     elif item == kWidgetID_withFrame:
#         withFrame = "{}".format(data == True)
#     elif item == kWidgetID_frameWidthSelector:
#         frameWidthSelector = vs.GetChoiceText(importDialog, kWidgetID_frameWidthSelector, data)
#     elif item == kWidgetID_frameHeightSelector:
#         frameHeightSelector = vs.GetChoiceText(importDialog, kWidgetID_frameHeightSelector, data)
#     elif item == kWidgetID_frameThicknessSelector:
#         vs.EnableItem(importDialog, kWidgetID_frameThickness, data == 0)
#         frameThicknessSelector = vs.GetChoiceText(importDialog, kWidgetID_frameThicknessSelector, data)
#     elif item == kWidgetID_frameThickness:
#         _, frameThickness = vs.GetEditReal(importDialog, kWidgetID_frameThickness, 3)
#     elif item == kWidgetID_frameDepthSelector:
#         vs.EnableItem(importDialog, kWidgetID_frameDepth, data == 0)
#         frameDepthSelector = vs.GetChoiceText(importDialog, kWidgetID_frameDepthSelector, data)
#     elif item == kWidgetID_frameDepth:
#         _, frameDepth = vs.GetEditReal(importDialog, kWidgetID_frameDepth, 3)
#     elif item == kWidgetID_frameClassSelector:
#         vs.EnableItem(importDialog, kWidgetID_frameClass, data == 0)
#         frameClassSelector = vs.GetChoiceText(importDialog, kWidgetID_frameClassSelector, data)
#     elif item == kWidgetID_frameClass:
#         index, frameClass = vs.GetSelectedChoiceInfo(importDialog, kWidgetID_frameClass, 0)
#     elif item == kWidgetID_frameTextureScaleSelector:
#         vs.EnableItem(importDialog, kWidgetID_frameTextureScale, data == 0)
#         frameTextureScaleSelector = vs.GetChoiceText(importDialog, kWidgetID_frameTextureScaleSelector, data)
#     elif item == kWidgetID_frameTextureScale:
#         _, frameTextureScale = vs.GetEditReal(importDialog, kWidgetID_frameTextureScale, 1)
#     elif item == kWidgetID_frameTextureRotationSelector:
#         vs.EnableItem(importDialog, kWidgetID_frameTextureRotation, data == 0)
#         frameTextureRotationSelector = vs.GetChoiceText(importDialog, kWidgetID_frameTextureRotationSelector, data)
#     elif item == kWidgetID_frameTextureRotation:
#         _, frameTextureRotation = vs.GetEditReal(importDialog, kWidgetID_frameTextureRotation, 1)
#     elif item == kWidgetID_withMatboardSelector:
#         vs.EnableItem(importDialog, kWidgetID_withMatboard, data == 0)
#         withMatboardSelector = vs.GetChoiceText(importDialog, kWidgetID_withMatboardSelector, data)
#     elif item == kWidgetID_withMatboard:
#         withMatboard = "{}".format(data == True)
#     elif item == kWidgetID_matboardPositionSelector:
#         vs.EnableItem(importDialog, kWidgetID_matboardPosition, data == 0)
#         matboardPositionSelector = vs.GetChoiceText(importDialog, kWidgetID_matboardPositionSelector, data)
#     elif item == kWidgetID_matboardPosition:
#         _, matboardPosition = vs.GetEditReal(importDialog, kWidgetID_matboardPosition, 3)
#     elif item == kWidgetID_matboardClassSelector:
#         vs.EnableItem(importDialog, kWidgetID_matboardClass, data == 0)
#         matboardClassSelector = vs.GetChoiceText(importDialog, kWidgetID_matboardClassSelector, data)
#     elif item == kWidgetID_matboardClass:
#         index, matboardClass = vs.GetSelectedChoiceInfo(importDialog, kWidgetID_matboardClass, 0)
#     elif item == kWidgetID_matboardTextureScaleSelector:
#         vs.EnableItem(importDialog, kWidgetID_matboardTextureScale, data == 0)
#         matboardTextureScaleSelector = vs.GetChoiceText(importDialog, kWidgetID_matboardTextureScaleSelector, data)
#     elif item == kWidgetID_matboardTextureScale:
#         _, matboardTextureScale = vs.GetEditReal(importDialog, kWidgetID_matboardTextureScale, 1)
#     elif item == kWidgetID_matboardTextureRotatSelector:
#         vs.EnableItem(importDialog, kWidgetID_matboardTextureRotat, data == 0)
#         matboardTextureRotatSelector = vs.GetChoiceText(importDialog, kWidgetID_matboardTextureRotatSelector, data)
#     elif item == kWidgetID_matboardTextureRotat:
#         _, matboardTextureRotat = vs.GetEditReal(importDialog, kWidgetID_matboardTextureRotat, 1)
#     elif item == kWidgetID_withGlassSelector:
#         vs.EnableItem(importDialog, kWidgetID_withGlass, data == 0)
#         withGlassSelector = vs.GetChoiceText(importDialog, kWidgetID_withGlassSelector, data)
#     elif item == kWidgetID_withGlass:
#         withGlass = "{}".format(data == True)
#     elif item == kWidgetID_glassPositionSelector:
#         vs.EnableItem(importDialog, kWidgetID_glassPosition, data == 0)
#         glassPositionSelector = vs.GetChoiceText(importDialog, kWidgetID_glassPositionSelector, data)
#     elif item == kWidgetID_glassPosition:
#         _, glassPosition = vs.GetEditReal(importDialog, kWidgetID_glassPosition, 3)
#     elif item == kWidgetID_glassClassSelector:
#         vs.EnableItem(importDialog, kWidgetID_glassClass, data == 0)
#         glassClassSelector = vs.GetChoiceText(importDialog, kWidgetID_glassClassSelector, data)
#     elif item == kWidgetID_glassClass:
#         index, glassClass = vs.GetSelectedChoiceInfo(importDialog, kWidgetID_glassClass, 0)
#     elif item == kWidgetID_excelCriteriaSelector:
#         vs.EnableItem(importDialog, kWidgetID_excelCriteriaValue, data != 0)
#         newExcelCriteriaSelector = vs.GetChoiceText(importDialog, kWidgetID_excelCriteriaSelector, data)
#         if newExcelCriteriaSelector != excelCriteriaSelector:
#             excelCriteriaSelector = newExcelCriteriaSelector
#             updateCriteriaValue(False)
#             if data != 0:
#                 updateCriteriaValue(True)
#             else:
#                 index = vs.GetChoiceIndex(importDialog, kWidgetID_excelCriteriaValue, excelCriteriaValue)
#                 if index == -1:
#                     vs.SelectChoice(importDialog, kWidgetID_excelCriteriaValue, 0, True);
#                     excelCriteriaValue = "Select a value ..."
#                 else:
#                     vs.SelectChoice(importDialog, kWidgetID_excelCriteriaValue, index, True);
#     elif item == kWidgetID_excelCriteriaValue:
#         excelCriteriaValue = vs.GetChoiceText(importDialog, kWidgetID_excelCriteriaValue, data)
#     elif item == kWidgetID_symbolCreateSymbol:
#         symbolCreateSymbol = "{}".format(data == True)
#         state =  vs.GetSelectedChoiceIndex(importDialog, kWidgetID_symbolFolderSelector, 0) == 0 and data == True
#         vs.EnableItem(importDialog, kWidgetID_symbolFolderSelector, data)
#         vs.EnableItem(importDialog, kWidgetID_symbolFolder, state)
#     elif item == kWidgetID_symbolFolderSelector:
#         vs.EnableItem(importDialog, kWidgetID_symbolFolder, data == 0)
#         symbolFolderSelector = vs.GetChoiceText(importDialog, kWidgetID_symbolFolderSelector, data)
#     elif item == kWidgetID_importIgnoreErrors:
#         importIgnoreErrors = "{}".format(data == True)
#         vs.ShowItem(importDialog, kWidgetID_importErrorCount, data != True)
#     elif item == kWidgetID_importIgnoreExisting:
#         importIgnoreExisting = "{}".format(data == True)
#     elif item == kWidgetID_importIgnoreUnmodified:
#         importIgnoreUnmodified = "{}".format(data == True)
#     elif item == kWidgetID_importButton:
#         importPictures()
#         vs.SetItemText(importDialog, kWidgetID_importNewCount, "New Pictures: {}".format(importNewCount))
#         vs.SetItemText(importDialog, kWidgetID_importUpdatedCount, "Updated Pictures: {}".format(importUpdatedCount))
#         vs.SetItemText(importDialog, kWidgetID_importDeletedCount, "Deleted Pictures: {}".format(importDeletedCount))
#         vs.SetItemText(importDialog, kWidgetID_importErrorCount, "Error Pictures: {}".format(importErrorCount))
#
#
#     if item  == kWidgetID_fileName or item == kWidgetID_fileBrowseButton or item == KDialogInitEvent:
#         connectionString = 'Driver={{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}};DBQ={};ReadOnly=1;'.format(excelFileName)
#         if database != 0:
#             database.close()
#             database = 0
#         try:
#             database = pyodbc.connect(connectionString, autocommit=True)
#         except:
#             vs.SetItemText(importDialog, kWidgetID_excelSheetNameLabel, "Invalid Excel file!")
#             vs.AlertCritical(connectionString, "Talk to Carlos")
#
#         if database:
#             cursor = database.cursor()
#             if cursor:
#                 for row in cursor.tables():
#                     vs.AddChoice(importDialog, kWidgetID_excelSheetName, row['table_name'], 0)
#                 cursor.close
#                 vs.AddChoice(importDialog, kWidgetID_excelSheetName, "Select an excel sheet", 0)
#                 index = vs.GetChoiceIndex(importDialog, kWidgetID_excelSheetName, excelSheetName)
#                 if index == -1:
#                     vs.SelectChoice(importDialog, kWidgetID_excelSheetName, 0, True);
#                     excelSheetName = "Select an excel sheet"
#                 else:
#                     vs.SelectChoice(importDialog, kWidgetID_excelSheetName, index, True);
#                     showParameters(True)
#
#                 vs.SetItemText(importDialog, kWidgetID_excelSheetNameLabel, "Excel sheet: ")
#                 vs.ShowItem(importDialog, kWidgetID_excelSheetNameLabel, True)
#                 vs.ShowItem(importDialog, kWidgetID_excelSheetName, True)
#         else:
#             while vs.GetChoiceCount(importDialog, kWidgetID_excelSheetName):
#                 vs.RemoveChoice(importDialog, kWidgetID_excelSheetName, 0)
#             vs.ShowItem(importDialog, kWidgetID_excelSheetNameLabel, True)
#             vs.ShowItem(importDialog, kWidgetID_excelSheetName, False)
#             showParameters(False)
#
#
#
#     if item == kWidgetID_withImageSelector or item == kWidgetID_withImage or item == kWidgetID_excelSheetName:
#         state = vs.GetSelectedChoiceIndex(importDialog, kWidgetID_withImageSelector, 0) != 0 or vs.GetBooleanItem(importDialog, kWidgetID_withImage) == True
#         vs.EnableItem(importDialog, kWidgetID_imageWidthLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_imageWidthSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_imageHeightLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_imageHeightSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_imagePositionLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_imagePositionSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_imagePosition, state)
#         vs.EnableItem(importDialog, kWidgetID_imageTextureLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_imageTextureSelector, state)
#
#     if item == kWidgetID_withFrameSelector or item == kWidgetID_withFrame or item == kWidgetID_excelSheetName:
#         state =  vs.GetSelectedChoiceIndex(importDialog, kWidgetID_withFrameSelector, 0) != 0 or vs.GetBooleanItem(importDialog, kWidgetID_withFrame) == True
#         vs.EnableItem(importDialog, kWidgetID_frameWidthLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_frameWidthSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_frameHeightLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_frameHeightSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_frameThicknessLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_frameThicknessSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_frameThickness, state)
#         vs.EnableItem(importDialog, kWidgetID_frameDepthLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_frameDepthSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_frameDepth, state)
#         vs.EnableItem(importDialog, kWidgetID_frameClassLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_frameClassSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_frameClass, state)
#         vs.EnableItem(importDialog, kWidgetID_frameTextureScaleLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_frameTextureScaleSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_frameTextureScale, state)
#         vs.EnableItem(importDialog, kWidgetID_frameTextureRotationLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_frameTextureRotationSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_frameTextureRotation, state)
#
#     if item == kWidgetID_withMatboardSelector or item == kWidgetID_withMatboard or item == kWidgetID_excelSheetName:
#         state =  vs.GetSelectedChoiceIndex(importDialog, kWidgetID_withMatboardSelector, 0) != 0 or vs.GetBooleanItem(importDialog, kWidgetID_withMatboard) == True
#         vs.EnableItem(importDialog, kWidgetID_matboardPositionLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardPositionSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardPosition, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardClassLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardClassSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardClass, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardTextureScaleLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardTextureScaleSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardTextureScale, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardTextureRotatLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardTextureRotatSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_matboardTextureRotat, state)
#
#     if item == kWidgetID_withGlassSelector or item == kWidgetID_withGlass or item == kWidgetID_excelSheetName:
#         state =  vs.GetSelectedChoiceIndex(importDialog, kWidgetID_withGlassSelector, 0) != 0 or vs.GetBooleanItem(importDialog, kWidgetID_withGlass) == True
#         vs.EnableItem(importDialog, kWidgetID_glassPositionLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_glassPositionSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_glassPosition, state)
#         vs.EnableItem(importDialog, kWidgetID_glassClassLabel, state)
#         vs.EnableItem(importDialog, kWidgetID_glassClassSelector, state)
#         vs.EnableItem(importDialog, kWidgetID_glassClass, state)
#
#     imageValid =    ((withImageSelector == "-- Manual" and withImage == "True") or withImageSelector != "-- Manual") and \
#                     (imageTextureSelector != "-- Select column ...") and \
#                     (imageWidthSelector != "-- Select column ...") and \
#                     (imageHeightSelector != "-- Select column ...")
#
#     frameValid =    ((withFrameSelector == "-- Manual" and withFrame == "True") or withFrameSelector != "-- Manual") and \
#                     (frameWidthSelector != "-- Select column ...") and \
#                     (frameHeightSelector != "-- Select column ...")
#
#     matboardValid = ((withMatboardSelector == "-- Manual" and withMatboard == "True") or withMatboardSelector != "-- Manual")
#
#     glassValid = ((withGlassSelector == "-- Manual" and withGlass == "True") or withGlassSelector != "-- Manual")
#
#     criteriaValid = (excelCriteriaSelector != "-- Select column ..." and excelCriteriaValue != "Select a value ..." )
#
#     importValid = (imageValid or frameValid ) and criteriaValid
#
#     vs.EnableItem(importDialog, kWidgetID_importButton, importValid)
#     vs.EnableItem(importDialog, kWidgetID_importNewCount, importValid)
#     vs.EnableItem(importDialog, kWidgetID_importUpdatedCount, importValid)
#     vs.EnableItem(importDialog, kWidgetID_importDeletedCount, importValid)
#
#
# def createImportDialog():
#
#     global importDialog
#     global excelFileName
#     global excelSheetName
#     global withImage
#     global imageTexure
#     global imageWidth
#     global imageHeight
#     global imagePosition
#     global withFrame
#     global frameWidth
#     global frameHeight
#     global frameThickness
#     global frameDepth
#     global frameClass
#     global frameTextureScale
#     global frameTextureRotation
#     global withMatboard
#     global matboardPosition
#     global matboardClass
#     global matboardTextureScale
#     global matboardTextureRotat
#     global withGlass
#     global glassPosition
#     global glassClass
#     global criteriaSelector
#     global criteriaValue
#
#     global withImageSelector
#     global imageTextureSelector
#     global imageWidthSelector
#     global imageHeightSelector
#     global imagePositionSelector
#     global withFrameSelector
#     global frameWidthSelector
#     global frameHeightSelector
#     global frameThicknessSelector
#     global frameDepthSelector
#     global frameClassSelector
#     global frameTextureScaleSelector
#     global frameTextureRotationSelector
#     global withMatboardSelector
#     global matboardPositionSelector
#     global matboardClassSelector
#     global matboardTextureScaleSelector
#     global matboardTextureRotatSelector
#     global withGlassSelector
#     global glassPositionSelector
#     global glassClassSelector
#     global criteriaSelector
#     global criteriaValue
#
#     global importIgnoreErrors
#     global importIgnoreExisting
#     global importIgnoreUnmodified
#     global symbolCreateSymbol
#     global symbolFolderSelector
#     global symbolFolder
#     global importNewCount
#     global importUpdatedCount
#     global importDeletedCount
#     global importErrorCount
#
#     inputFieldWidth = 20
#     labelWidth = 20
#
#     importNewCount      = 0
#     importUpdatedCount  = 0
#     importDeletedCount  = 0
#     importErrorCount    = 0
#
#
#     importDialog = vs.CreateLayout("Import Pictures", True, "OK", "Cancel")
#
#     # Excel file group
#     # =========================================================================================
#     vs.CreateGroupBox(importDialog, kWidgetID_excelFileGroup, "Excel spreadsheet", True)
#     vs.SetFirstLayoutItem(importDialog, kWidgetID_excelFileGroup)
#     # File Name
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_fileNameLabel, "Excel file: ", -1)
#     vs.SetFirstGroupItem(importDialog, kWidgetID_excelFileGroup, kWidgetID_fileNameLabel)
#     vs.CreateEditText (importDialog, kWidgetID_fileName, excelFileName, 3 * inputFieldWidth )
#     vs.SetRightItem(importDialog, kWidgetID_fileNameLabel, kWidgetID_fileName, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_fileName, "Enter the excel file name here")
#     # File browse button
#     # -----------------------------------------------------------------------------------------
#     vs.CreatePushButton(importDialog, kWidgetID_fileBrowseButton, "Browse...")
#     vs.SetRightItem(importDialog, kWidgetID_fileName, kWidgetID_fileBrowseButton, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_fileBrowseButton, "Click to browse Excel file")
#     # Excel sheet selection
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_excelSheetNameLabel, "Excel sheet: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_fileNameLabel, kWidgetID_excelSheetNameLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_excelSheetName, inputFieldWidth)
#     sheetIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_excelSheetName, excelSheetName)
#     vs.SelectChoice(importDialog, kWidgetID_excelSheetName, sheetIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_excelSheetNameLabel, kWidgetID_excelSheetName, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_excelSheetName, "Select the Excel sheet")
#
#     # Image group
#     # =========================================================================================
#     vs.CreateGroupBox(importDialog, kWidgetID_imageGroup, "Image", True)
#     vs.SetBelowItem(importDialog, kWidgetID_excelFileGroup, kWidgetID_imageGroup, 0, 0)
#     # With Image checkbox
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_withImageLabel, "With Image: ", labelWidth)
#     vs.SetFirstGroupItem(importDialog, kWidgetID_imageGroup, kWidgetID_withImageLabel)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_withImageSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_withImageSelector, withImageSelector)
#     vs.SelectChoice(importDialog, kWidgetID_withImageSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_withImageLabel, kWidgetID_withImageSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_withImageSelector, "Select the column for the image creation")
#     vs.CreateCheckBox (importDialog, kWidgetID_withImage, "Include Image")
#     vs.SetBooleanItem(importDialog, kWidgetID_withImage, withImage == "True")
#     vs.SetRightItem(importDialog, kWidgetID_withImageSelector, kWidgetID_withImage, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_withImage, "Choose the value for the image creation")
#     # Image Folder Name
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_imageFolderNameLabel, "Images folder: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_withImageLabel, kWidgetID_imageFolderNameLabel, 0, 0)
#     vs.CreateEditText (importDialog, kWidgetID_imageFolderName, imageFolderName, inputFieldWidth )
#     vs.SetRightItem(importDialog, kWidgetID_imageFolderNameLabel, kWidgetID_imageFolderName, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_imageFolderName, "Enter the folder for the image files")
#     # File browse button
#     # -----------------------------------------------------------------------------------------
#     vs.CreatePushButton(importDialog, kWidgetID_imageFolderBrowseButton, "Browse...")
#     vs.SetRightItem(importDialog, kWidgetID_imageFolderName, kWidgetID_imageFolderBrowseButton, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_imageFolderBrowseButton, "Click to browse the images folder")
#     # Image Texture
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_imageTextureLabel, "Image name: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_imageFolderNameLabel, kWidgetID_imageTextureLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_imageTextureSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_imageTextureSelector, imageTextureSelector)
#     vs.SelectChoice(importDialog, kWidgetID_imageTextureSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_imageTextureLabel, kWidgetID_imageTextureSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_imageTextureSelector, "Select the column for the image name")
#     # Image Width dimension
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_imageWidthLabel, "Image Width: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_imageTextureLabel, kWidgetID_imageWidthLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_imageWidthSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_imageWidthSelector, imageWidthSelector)
#     vs.SelectChoice(importDialog, kWidgetID_imageWidthSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_imageWidthLabel, kWidgetID_imageWidthSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_imageWidthSelector, "Select the column for the image width")
#     # Image Height dimension
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_imageHeightLabel, "Image Height: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_imageWidthLabel, kWidgetID_imageHeightLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_imageHeightSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_imageHeightSelector, imageHeightSelector)
#     vs.SelectChoice(importDialog, kWidgetID_imageHeightSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_imageHeightLabel, kWidgetID_imageHeightSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_imageHeightSelector, "Select the column for the image height")
#     # Image Position dimension
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_imagePositionLabel, "Image Position: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_imageHeightLabel, kWidgetID_imagePositionLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_imagePositionSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_imagePositionSelector, imagePositionSelector)
#     vs.SelectChoice(importDialog, kWidgetID_imagePositionSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_imagePositionLabel, kWidgetID_imagePositionSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_imagePositionSelector, "Select the column for the image position")
#     vs.CreateEditReal(importDialog, kWidgetID_imagePosition, 3, imagePosition, inputFieldWidth)
#     vs.SetRightItem(importDialog, kWidgetID_imagePositionSelector, kWidgetID_imagePosition, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_imagePosition, "Enter the position (depth) of the image here.")
#
#     # Frame group
#     # =========================================================================================
#     vs.CreateGroupBox(importDialog, kWidgetID_frameGroup, "Frame", True)
#     vs.SetBelowItem(importDialog, kWidgetID_imageGroup, kWidgetID_frameGroup, 0, 0)
#     # With Frame checkbox
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_withFrameLabel, "With Frame: ", labelWidth)
#     vs.SetFirstGroupItem(importDialog, kWidgetID_frameGroup, kWidgetID_withFrameLabel)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_withFrameSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_withFrameSelector, withFrameSelector)
#     vs.SelectChoice(importDialog, kWidgetID_withFrameSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_withFrameLabel, kWidgetID_withFrameSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_withFrameSelector, "Select the column for the frame creation")
#     vs.CreateCheckBox (importDialog, kWidgetID_withFrame, "Include Frame")
#     vs.SetBooleanItem(importDialog, kWidgetID_withFrame, withImage == "True")
#     vs.SetRightItem(importDialog, kWidgetID_withFrameSelector, kWidgetID_withFrame, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_withFrame, "Choose the value for the frame creation")
#     # Frame Width dimension
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_frameWidthLabel, "Frame Width: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_withFrameLabel, kWidgetID_frameWidthLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_frameWidthSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameWidthSelector, frameWidthSelector)
#     vs.SelectChoice(importDialog, kWidgetID_frameWidthSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_frameWidthLabel, kWidgetID_frameWidthSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameWidthSelector, "Select the column for the frame width")
#     # Frame Height dimension
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_frameHeightLabel, "Frame Height: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_frameWidthLabel, kWidgetID_frameHeightLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_frameHeightSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameHeightSelector, frameHeightSelector)
#     vs.SelectChoice(importDialog, kWidgetID_frameHeightSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_frameHeightLabel, kWidgetID_frameHeightSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameHeightSelector, "Select the column for the frame height")
#     # Frame Thickness dimension
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_frameThicknessLabel, "Frame Thickness: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_frameHeightLabel, kWidgetID_frameThicknessLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_frameThicknessSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameThicknessSelector, frameThicknessSelector)
#     vs.SelectChoice(importDialog, kWidgetID_frameThicknessSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_frameThicknessLabel, kWidgetID_frameThicknessSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameThicknessSelector, "Select the column for the frame thickness")
#     vs.CreateEditReal(importDialog, kWidgetID_frameThickness, 3, frameThickness, inputFieldWidth)
#     vs.SetRightItem(importDialog, kWidgetID_frameThicknessSelector, kWidgetID_frameThickness, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameThickness, "Enter the thickness of the frame here.")
#     # Frame Depth dimension
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_frameDepthLabel, "Frame Depth: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_frameThicknessLabel, kWidgetID_frameDepthLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_frameDepthSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameDepthSelector, frameDepthSelector)
#     vs.SelectChoice(importDialog, kWidgetID_frameDepthSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_frameDepthLabel, kWidgetID_frameDepthSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameDepthSelector, "Select the column for the frame depth")
#     vs.CreateEditReal(importDialog, kWidgetID_frameDepth, 3, frameDepth, inputFieldWidth)
#     vs.SetRightItem(importDialog, kWidgetID_frameDepthSelector, kWidgetID_frameDepth, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameDepth, "Enter the depth of the frame here.")
#     # Frame Class
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_frameClassLabel, "Frame Class: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_frameDepthLabel, kWidgetID_frameClassLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_frameClassSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameClassSelector, frameClassSelector)
#     vs.SelectChoice(importDialog, kWidgetID_frameClassSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_frameClassLabel, kWidgetID_frameClassSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameClassSelector, "Select the column for the frame class")
#     vs.CreateClassPullDownMenu(importDialog, kWidgetID_frameClass, inputFieldWidth)
#     classIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameClass, frameClass)
#     vs.SelectChoice(importDialog, kWidgetID_frameClass, classIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_frameClassSelector, kWidgetID_frameClass, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameClass, "Enter the class of the frame here.")
#     # Frame Texture scale
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_frameTextureScaleLabel, "Frame Texture Scale: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_frameClassLabel, kWidgetID_frameTextureScaleLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_frameTextureScaleSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameTextureScaleSelector, frameTextureScaleSelector)
#     vs.SelectChoice(importDialog, kWidgetID_frameTextureScaleSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_frameTextureScaleLabel, kWidgetID_frameTextureScaleSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameTextureScaleSelector, "Select the column for the frame texture scale")
#     vs.CreateEditReal(importDialog, kWidgetID_frameTextureScale, 1, frameTextureScale, inputFieldWidth)
#     vs.SetRightItem(importDialog, kWidgetID_frameTextureScaleSelector, kWidgetID_frameTextureScale, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameTextureScale, "Enter the frame texture scale")
#     # Frame Texture rotation
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_frameTextureRotationLabel, "Frame Texture Rotation: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_frameTextureScaleLabel, kWidgetID_frameTextureRotationLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_frameTextureRotationSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_frameTextureRotationSelector, frameTextureRotationSelector)
#     vs.SelectChoice(importDialog, kWidgetID_frameTextureRotationSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_frameTextureRotationLabel, kWidgetID_frameTextureRotationSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameTextureRotationSelector, "Select the column for the frame texture rotation")
#     vs.CreateEditReal(importDialog, kWidgetID_frameTextureRotation, 1, frameTextureRotation, inputFieldWidth)
#     vs.SetRightItem(importDialog, kWidgetID_frameTextureRotationSelector, kWidgetID_frameTextureRotation, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_frameTextureRotation, "Enter the frame texture scale")
#
#     # Matboard group
#     # =========================================================================================
#     vs.CreateGroupBox(importDialog, kWidgetID_matboardGroup, "Matboard", True)
#     vs.SetBelowItem(importDialog, kWidgetID_frameGroup, kWidgetID_matboardGroup, 0, 0)
#
#     # With Matboard checkbox
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_withMatboardLabel, "With Matboard: ", labelWidth)
#     vs.SetFirstGroupItem(importDialog, kWidgetID_matboardGroup, kWidgetID_withMatboardLabel)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_withMatboardSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_withMatboardSelector, withMatboardSelector)
#     vs.SelectChoice(importDialog, kWidgetID_withMatboardSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_withMatboardLabel, kWidgetID_withMatboardSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_withMatboardSelector, "Select the column for the Matboard creation")
#     vs.CreateCheckBox (importDialog, kWidgetID_withMatboard, "Include Matboard")
#     vs.SetBooleanItem(importDialog, kWidgetID_withMatboard, withMatboard == "True")
#     vs.SetRightItem(importDialog, kWidgetID_withMatboardSelector, kWidgetID_withMatboard, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_withMatboard, "Choose the value for the Matboard creation")
#
#     # Matboard Position dimension
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_matboardPositionLabel, "Matboard Position: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_withMatboardLabel, kWidgetID_matboardPositionLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_matboardPositionSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_matboardPositionSelector, matboardPositionSelector)
#     vs.SelectChoice(importDialog, kWidgetID_matboardPositionSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_matboardPositionLabel, kWidgetID_matboardPositionSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_matboardPositionSelector, "Select the column for the matboard position")
#     vs.CreateEditReal(importDialog, kWidgetID_matboardPosition, 3, matboardPosition, inputFieldWidth)
#     vs.SetRightItem(importDialog, kWidgetID_matboardPositionSelector, kWidgetID_matboardPosition, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_matboardPosition, "Enter the position (depth) of the matboard here.")
#     # Matboard Class
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_matboardClassLabel, "Matboard Class: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_matboardPositionLabel, kWidgetID_matboardClassLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_matboardClassSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_matboardClassSelector, matboardClassSelector)
#     vs.SelectChoice(importDialog, kWidgetID_matboardClassSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_matboardClassLabel, kWidgetID_matboardClassSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_matboardClassSelector, "Select the column for the matboard class")
#     vs.CreateClassPullDownMenu(importDialog, kWidgetID_matboardClass, inputFieldWidth)
#     classIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_matboardClass, matboardClass)
#     vs.SelectChoice(importDialog, kWidgetID_matboardClass, classIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_matboardClassSelector, kWidgetID_matboardClass, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_matboardClass, "Enter the class of the matboard here.")
#     # Frame Texture scale
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_matboardTextureScaleLabel, "Matboard Texture Scale: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_matboardClassLabel, kWidgetID_matboardTextureScaleLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_matboardTextureScaleSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_matboardTextureScaleSelector, matboardTextureScaleSelector)
#     vs.SelectChoice(importDialog, kWidgetID_matboardTextureScaleSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_matboardTextureScaleLabel, kWidgetID_matboardTextureScaleSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_matboardTextureScaleSelector, "Select the column for the matboard texture scale")
#     vs.CreateEditReal(importDialog, kWidgetID_matboardTextureScale, 1, matboardTextureScale, inputFieldWidth)
#     vs.SetRightItem(importDialog, kWidgetID_matboardTextureScaleSelector, kWidgetID_matboardTextureScale, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_matboardTextureScale, "Enter the matboard texture scale")
#     # Frame Texture rotation
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_matboardTextureRotatLabel, "Matboard Texture Rotation: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_matboardTextureScaleLabel, kWidgetID_matboardTextureRotatLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_matboardTextureRotatSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_matboardTextureRotatSelector, matboardTextureRotatSelector)
#     vs.SelectChoice(importDialog, kWidgetID_matboardTextureRotatSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_matboardTextureRotatLabel, kWidgetID_matboardTextureRotatSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_matboardTextureRotatSelector, "Select the column for the matboard texture rotation")
#     vs.CreateEditReal(importDialog, kWidgetID_matboardTextureRotat, 1, matboardTextureRotat, inputFieldWidth)
#     vs.SetRightItem(importDialog, kWidgetID_matboardTextureRotatSelector, kWidgetID_matboardTextureRotat, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_matboardTextureRotat, "Enter the matboard texture scale")
#
#     # Glass group
#     # =========================================================================================
#     vs.CreateGroupBox(importDialog, kWidgetID_glassGroup, "Glass", True)
#     vs.SetBelowItem(importDialog, kWidgetID_matboardGroup, kWidgetID_glassGroup, 0, 0)
#
#     # With Glass checkbox
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_withGlassLabel, "With Glass: ", labelWidth)
#     vs.SetFirstGroupItem(importDialog, kWidgetID_glassGroup, kWidgetID_withGlassLabel)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_withGlassSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_withGlassSelector, withGlassSelector)
#     vs.SelectChoice(importDialog, kWidgetID_withGlassSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_withGlassLabel, kWidgetID_withGlassSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_withGlassSelector, "Select the column for the Glass creation")
#     vs.CreateCheckBox (importDialog, kWidgetID_withGlass, "Include Galss")
#     vs.SetBooleanItem(importDialog, kWidgetID_withGlass, withGlass == "True")
#     vs.SetRightItem(importDialog, kWidgetID_withGlassSelector, kWidgetID_withGlass, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_withGlass, "Choose the value for the Glass creation")
#     # Glass Position dimension
#     # -----------------------------------------------------------------------------------------
#
#     vs.CreateStaticText(importDialog, kWidgetID_glassPositionLabel, "Glass Position: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_withGlassLabel, kWidgetID_glassPositionLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_glassPositionSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_glassPositionSelector, glassPositionSelector)
#     vs.SelectChoice(importDialog, kWidgetID_glassPositionSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_glassPositionLabel, kWidgetID_glassPositionSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_glassPositionSelector, "Select the column for the glass position")
#     vs.CreateEditReal(importDialog, kWidgetID_glassPosition, 3, glassPosition, inputFieldWidth)
#     vs.SetRightItem(importDialog, kWidgetID_glassPositionSelector, kWidgetID_glassPosition, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_glassPosition, "Enter the position (depth) of the glass here.")
#     # Glass Class
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_glassClassLabel, "Glass Class: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_glassPositionLabel, kWidgetID_glassClassLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_glassClassSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_glassClassSelector, glassClassSelector)
#     vs.SelectChoice(importDialog, kWidgetID_glassClassSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_glassClassLabel, kWidgetID_glassClassSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_glassClassSelector, "Select the column for the glass class")
#     vs.CreateClassPullDownMenu(importDialog, kWidgetID_glassClass, inputFieldWidth)
#     classIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_glassClass, glassClass)
#     vs.SelectChoice(importDialog, kWidgetID_glassClass, classIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_glassClassSelector, kWidgetID_glassClass, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_glassClass, "Enter the class of the glass here.")
#
#     # Criteria group
#     # =========================================================================================
#     vs.CreateGroupBox(importDialog, kWidgetID_excelCriteriaGroup, "Criteria", True)
#     vs.SetRightItem(importDialog, kWidgetID_imageGroup, kWidgetID_excelCriteriaGroup, 0, 0)
#     # Criteria
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_excelCriteriaLabel, "Picture Creation Criteria: ", labelWidth)
#     vs.SetFirstGroupItem(importDialog, kWidgetID_excelCriteriaGroup, kWidgetID_excelCriteriaLabel)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_excelCriteriaSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_excelCriteriaSelector, excelCriteriaSelector)
#     vs.SelectChoice(importDialog, kWidgetID_excelCriteriaSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_excelCriteriaLabel, kWidgetID_excelCriteriaSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_excelCriteriaSelector, "Select the column for selection criteria")
#
#     vs.CreatePullDownMenu(importDialog, kWidgetID_excelCriteriaValue, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_excelCriteriaValue, excelCriteriaValue)
#     vs.SelectChoice(importDialog, kWidgetID_excelCriteriaValue, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_excelCriteriaSelector, kWidgetID_excelCriteriaValue, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_excelCriteriaValue, "Select the selection criteria value")
#
#     # Symbol group
#     # =========================================================================================
#     vs.CreateGroupBox(importDialog, kWidgetID_symbolGroup, "Symbol", True)
#     vs.SetBelowItem(importDialog, kWidgetID_excelCriteriaGroup, kWidgetID_symbolGroup, 0, 0)
#     # Create Symbol checkbox
#     # -----------------------------------------------------------------------------------------
#     vs.CreateCheckBox(importDialog, kWidgetID_symbolCreateSymbol, "Create Symbol")
#     vs.SetFirstGroupItem(importDialog, kWidgetID_symbolGroup, kWidgetID_symbolCreateSymbol)
#     vs.SetBooleanItem(importDialog, kWidgetID_symbolCreateSymbol, symbolCreateSymbol == "True")
#     vs.SetHelpText(importDialog, kWidgetID_symbolCreateSymbol, "Check to create a symbol for every Picture")
#     # Symbol Folder
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_symbolFolderLabel, "Symbol Folder: ", labelWidth)
#     vs.SetBelowItem(importDialog, kWidgetID_symbolCreateSymbol, kWidgetID_symbolFolderLabel, 0, 0)
#     vs.CreatePullDownMenu(importDialog, kWidgetID_symbolFolderSelector, inputFieldWidth)
#     selectorIndex = vs.GetPopUpChoiceIndex(importDialog, kWidgetID_symbolFolderSelector, symbolFolderSelector)
#     vs.SelectChoice(importDialog, kWidgetID_symbolFolderSelector, selectorIndex, True)
#     vs.SetRightItem(importDialog, kWidgetID_symbolFolderLabel, kWidgetID_symbolFolderSelector, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_symbolFolderSelector, "Select the column for the symbol folder name")
#
#     vs.CreateEditText (importDialog, kWidgetID_symbolFolder, symbolFolder, inputFieldWidth )
#     vs.SetRightItem(importDialog, kWidgetID_symbolFolderSelector, kWidgetID_symbolFolder, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_symbolFolder, "Enter the symbol folder name")
#
#     # Import group
#     # =========================================================================================
#     vs.CreateGroupBox(importDialog, kWidgetID_importGroup, "Import", True)
#     vs.SetBelowItem(importDialog, kWidgetID_symbolGroup, kWidgetID_importGroup, 0, 0)
#     # Ignore Existing
#     # -----------------------------------------------------------------------------------------
#     vs.CreateCheckBox(importDialog, kWidgetID_importIgnoreExisting, "Ignore manual fields on existing Pictures")
#     vs.SetFirstGroupItem(importDialog, kWidgetID_importGroup, kWidgetID_importIgnoreExisting)
#     vs.SetBooleanItem(importDialog, kWidgetID_importIgnoreExisting, importIgnoreExisting == "True")
#     vs.SetHelpText(importDialog, kWidgetID_importIgnoreExisting, "Ignore manual fields on existing Pictures")
#     # Ignore Errors
#     # -----------------------------------------------------------------------------------------
#     vs.CreateCheckBox(importDialog, kWidgetID_importIgnoreErrors, "Ignore Errors")
#     vs.SetBelowItem(importDialog, kWidgetID_importIgnoreExisting, kWidgetID_importIgnoreErrors, 0, 0)
#     vs.SetBooleanItem(importDialog, kWidgetID_importIgnoreErrors, importIgnoreErrors == "True")
#     vs.SetHelpText(importDialog, kWidgetID_importIgnoreErrors, "Check to ignore all import errors")
#     # Ignore Unmodified
#     # -----------------------------------------------------------------------------------------
#     vs.CreateCheckBox(importDialog, kWidgetID_importIgnoreUnmodified, "Ignore Unmodified")
#     vs.SetBelowItem(importDialog, kWidgetID_importIgnoreErrors, kWidgetID_importIgnoreUnmodified, 0, 0)
#     vs.SetBooleanItem(importDialog, kWidgetID_importIgnoreUnmodified, importIgnoreUnmodified == "True")
#     vs.SetHelpText(importDialog, kWidgetID_importIgnoreUnmodified, "Check to ignore all unmodified pictures")
#
#     # Import Button
#     # -----------------------------------------------------------------------------------------
#     vs.CreatePushButton(importDialog, kWidgetID_importButton, "Import")
#     vs.SetBelowItem(importDialog, kWidgetID_importIgnoreUnmodified, kWidgetID_importButton, 0, 0)
#     vs.SetHelpText(importDialog, kWidgetID_fileBrowseButton, "Click to start the import operation")
#     # New Pictures Count
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_importNewCount, "New Pictures: {}".format(importNewCount), labelWidth + 10)
#     vs.SetBelowItem(importDialog, kWidgetID_importButton, kWidgetID_importNewCount, 0, 0)
#     # Updated Pictures Count
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_importUpdatedCount, "Updated Pictures: {}".format(importUpdatedCount), labelWidth + 10)
#     vs.SetBelowItem(importDialog, kWidgetID_importNewCount, kWidgetID_importUpdatedCount, 0, 0)
#     # Deleted Pictures Count
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_importDeletedCount, "Deleted Pictures: {}".format(importDeletedCount), labelWidth + 10)
#     vs.SetBelowItem(importDialog, kWidgetID_importUpdatedCount, kWidgetID_importDeletedCount, 0, 0)
#     # Error Pictures Count
#     # -----------------------------------------------------------------------------------------
#     vs.CreateStaticText(importDialog, kWidgetID_importErrorCount, "Error Pictures: {}".format(importErrorCount), labelWidth + 10)
#     vs.SetBelowItem(importDialog, kWidgetID_importDeletedCount, kWidgetID_importErrorCount, 0, 0)
#
#     return importDialog
#
# def pyODBCAccess():
#     #    importPt = (0,0)
#     baseDir = "E:\Documents\wdfm\Pinocchio\Object List"
#     excelFileName = baseDir + "\Pinocchio Object List_03.07.16.xlsm"
#     pictureName = "New Picture"
#     withImage = "True"
#     imageWidth = "10\""
#     imageHeight = "6\""
#     imagePosition = "0.3\""
#     withFrame = "True"
#     frameWidth = "8\""
#     frameHeight = "12\""
#     frameThickness = "1\""
#     frameDepth = "1\""
#     frameClass = "None"
#     frameTextureScale = "0.1\""
#     frameTextureRotation = "0\""
#     withMatboard = "True"
#     matboardPosition = "0.25\""
#     matboardClass = "None"
#     matboardTextureScale = "0.1\""
#     matboardTextureRotat = "0"
#     withGlass = "True"
#     glassPosition = "0.75"
#     glassClass = "None"
#     connectionString = 'Driver={{Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)}};DBQ={};ReadOnly=1;'.format(excelFileName)
#     queryString = 'SELECT * \
#                    FROM [Objects$] \
#                    WHERE [Type] = \'Object\' \
#                    AND [Artwork _Dimensions] IS NOT NULL \
#                    AND [F13] IS NOT NULL \
#                    AND [Frame/Mounting Dimensions] IS NOT NULL \
#                    AND [F16] IS NOT NULL;'
#
#     database = pyodbc.connect(connectionString, autocommit=True)
#     if database:
#         cursor = database.cursor()
#         if cursor:
#             for row in cursor.tables():
#                 tables = row['table_name']
#             cursor.close
#
#         cursor = database.cursor()
#         if cursor:
#             for row in cursor.columns('Objects$'):
#                 columns = row['column_name']
#             cursor.close
#
#         cursor = database.cursor()
#         if cursor:
#             i = 0
#             for row in cursor.execute(queryString):
#
#                 pictureName = ""
#                 withImage = "False"
#                 imageWidth = "0"
#                 imageHeight = "0"
#                 imagePosition = "0"
#                 withFrame = "False"
#                 frameWidth = "0"
#                 frameHeight = "0"
#                 frameThickness = "1"
#                 frameDepth = "1"
#                 frameClass = "Picture-Frame"
#                 frameTextureScale = "0.1"
#                 frameTextureRotation = "0"
#                 withMatboard = "True"
#                 matboardPosition = "0"
#                 matboardClass = "Picture-Matboard"
#                 matboardTextureScale = "0.1"
#                 matboardTextureRotat = "0"
#                 withGlass = "False"
#                 glassPosition = "0"
#                 glassClass = "Picture-Glass"
#
#                 directory = row["Room Location".lower()]
#                 pictureName = row["Image Name".lower()]
#                 imageHeight = row['Artwork _Dimensions'.lower()]
#                 imageWidth = row['F13'.lower()]
#                 frameHeight = row['Frame/Mounting Dimensions'.lower()]
#                 frameWidth = row['F16'.lower()]
#
#                 if pictureName != "":
#                     withImage = "True"
#                     if imageWidth != "" and imageHeight != "":
#                         withImage = "True"
#                     if frameWidth != "" and frameHeight != "":
#                         withFrame = "True"
#                         imagePosition = "{}".format(float(frameDepth) * 0.3)
#                         matboardPosition = "{}".format(float(frameDepth) * 0.25)
#                         glassPosition = "{}".format(float(frameDepth) * 0.75)
#                         updatePicture(  directory,
#                                         pictureName,
#                                         withImage,
#                                         imageWidth,
#                                         imageHeight,
#                                         imagePosition,
#                                         withFrame,
#                                         frameWidth,
#                                         frameHeight,
#                                         frameThickness,
#                                         frameDepth,
#                                         frameClass,
#                                         frameTextureScale,
#                                         frameTextureRotation,
#                                         withMatboard,
#                                         matboardPosition,
#                                         matboardClass,
#                                         matboardTextureScale,
#                                         matboardTextureRotat,
#                                         withGlass,
#                                         glassPosition,
#                                         glassClass)
#                 i = i + 1
#                 if i > 4: break
#             cursor.close
#         database.close
#

def execute():
    settings = ImportSettings()
    import_dialog = ImportPicturesDialog(settings)
    if import_dialog.result == kOK:
        settings.save()

# import_dialog = createImportDialog()
    # if vs.RunLayoutDialog(import_dialog, importDialogHandler) == kOK:
    #     settings.save()
