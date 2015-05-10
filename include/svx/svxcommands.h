/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*
 * This file is part of the LibreOffice project.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * This file incorporates work covered by the following license notice:
 *
 *   Licensed to the Apache Software Foundation (ASF) under one or more
 *   contributor license agreements. See the NOTICE file distributed
 *   with this work for additional information regarding copyright
 *   ownership. The ASF licenses this file to you under the Apache
 *   License, Version 2.0 (the "License"); you may not use this file
 *   except in compliance with the License. You may obtain a copy of
 *   the License at http://www.apache.org/licenses/LICENSE-2.0 .
 */
#ifndef INCLUDED_SVX_SVXCOMMANDS_H
#define INCLUDED_SVX_SVXCOMMANDS_H

#define CMD_SID_OBJECT_ALIGN_CENTER                 ".uno:AlignCenter"
#define CMD_SID_OBJECT_ALIGN_DOWN                   ".uno:AlignDown"
#define CMD_SID_OBJECT_ALIGN_LEFT                   ".uno:ObjectAlignLeft"
#define CMD_SID_OBJECT_ALIGN_MIDDLE                 ".uno:AlignMiddle"
#define CMD_SID_OBJECT_ALIGN_RIGHT                  ".uno:ObjectAlignRight"
#define CMD_SID_OBJECT_ALIGN_UP                     ".uno:AlignUp"
#define CMD_SID_FM_AUTOCONTROLFOCUS                 ".uno:AutoControlFocus"
#define CMD_SID_BEZIER_CLOSE                        ".uno:BezierClose"
#define CMD_SID_BEZIER_CONVERT                      ".uno:BezierConvert"
#define CMD_SID_BEZIER_CUTLINE                      ".uno:BezierCutLine"
#define CMD_SID_BEZIER_DELETE                       ".uno:BezierDelete"
#define CMD_SID_BEZIER_EDGE                         ".uno:BezierEdge"
#define CMD_SID_BEZIER_ELIMINATE_POINTS             ".uno:BezierEliminatePoints"
#define CMD_SID_BEZIER_INSERT                       ".uno:BezierInsert"
#define CMD_SID_BEZIER_MOVE                         ".uno:BezierMove"
#define CMD_SID_BEZIER_SMOOTH                       ".uno:BezierSmooth"
#define CMD_SID_BEZIER_SYMMTR                       ".uno:BezierSymmetric"
#define CMD_SID_ATTR_CHAR_WEIGHT                    ".uno:Bold"
#define CMD_SID_FRAME_TO_TOP                        ".uno:BringToFront"
#define CMD_SID_ATTR_PARA_ADJUST_CENTER             ".uno:CenterPara"
#define CMD_SID_FM_CHANGECONTROLTYPE                ".uno:ChangeControlType"
#define CMD_SID_ATTR_CHAR_FONT                      ".uno:CharFontName"
#define CMD_SID_FM_CHECKBOX                         ".uno:CheckBox"
#define CMD_SID_FM_COMBOBOX                         ".uno:ComboBox"
#define CMD_SID_CONTOUR_DLG                         ".uno:ContourDialog"
#define CMD_SID_FM_CONVERTTO_BUTTON                 ".uno:ConvertToButton"
#define CMD_SID_FM_CONVERTTO_CHECKBOX               ".uno:ConvertToCheckBox"
#define CMD_SID_FM_CONVERTTO_COMBOBOX               ".uno:ConvertToCombo"
#define CMD_SID_FM_CONVERTTO_CURRENCY               ".uno:ConvertToCurrency"
#define CMD_SID_FM_CONVERTTO_DATE                   ".uno:ConvertToDate"
#define CMD_SID_FM_CONVERTTO_EDIT                   ".uno:ConvertToEdit"
#define CMD_SID_FM_CONVERTTO_FILECONTROL            ".uno:ConvertToFileControl"
#define CMD_SID_FM_CONVERTTO_FIXEDTEXT              ".uno:ConvertToFixed"
#define CMD_SID_FM_CONVERTTO_FORMATTED              ".uno:ConvertToFormatted"
#define CMD_SID_FM_CONVERTTO_SCROLLBAR              ".uno:ConvertToScrollBar"
#define CMD_SID_FM_CONVERTTO_SPINBUTTON             ".uno:ConvertToSpinButton"
#define CMD_SID_FM_CONVERTTO_GROUPBOX               ".uno:ConvertToGroup"
#define CMD_SID_FM_CONVERTTO_IMAGEBUTTON            ".uno:ConvertToImageBtn"
#define CMD_SID_FM_CONVERTTO_IMAGECONTROL           ".uno:ConvertToImageControl"
#define CMD_SID_FM_CONVERTTO_LISTBOX                ".uno:ConvertToList"
#define CMD_SID_FM_CONVERTTO_NUMERIC                ".uno:ConvertToNumeric"
#define CMD_SID_FM_CONVERTTO_PATTERN                ".uno:ConvertToPattern"
#define CMD_SID_FM_CONVERTTO_RADIOBUTTON            ".uno:ConvertToRadio"
#define CMD_SID_FM_CONVERTTO_TIME                   ".uno:ConvertToTime"
#define CMD_SID_FM_CONVERTTO_NAVIGATIONBAR          ".uno:ConvertToNavigationBar"
#define CMD_SID_FM_CURRENCYFIELD                    ".uno:CurrencyField"
#define CMD_SID_FM_DATEFIELD                        ".uno:DateField"
#define CMD_SID_DISTRIBUTE_DLG                      ".uno:DistributeSelection"
#define CMD_SID_FM_EDIT                             ".uno:Edit"
#define CMD_SID_ENTER_GROUP                         ".uno:EnterGroup"
#define CMD_SID_CHAR_DLG                            ".uno:FontDialog"
#define CMD_SID_ATTR_CHAR_FONTHEIGHT                ".uno:FontHeight"
#define CMD_SID_FONTWORK                            ".uno:FontWork"
#define CMD_SID_ATTRIBUTES_AREA                     ".uno:FormatArea"
#define CMD_SID_GROUP                               ".uno:FormatGroup"
#define CMD_SID_ATTRIBUTES_LINE                     ".uno:FormatLine"
#define CMD_SID_FM_FORMATTEDFIELD                   ".uno:FormattedField"
#define CMD_SID_UNGROUP                             ".uno:FormatUngroup"
#define CMD_SID_GALLERY_ENABLE_ADDCOPY              ".uno:GalleryEnableAddCopy"
#define CMD_SID_GALLERY_FORMATS                     ".uno:InsertGalleryPic"
#define CMD_SID_GRID_USE                            ".uno:GridUse"
#define CMD_SID_GRID_VISIBLE                        ".uno:GridVisible"
#define CMD_SID_INSERT_POSTIT                       ".uno:InsertAnnotation"
#define CMD_SID_REPLYTO_POSTIT                      ".uno:ReplyToAnnotation"
#define CMD_SID_RULER                               ".uno:ShowRuler"
#define CMD_SID_DELETE_POSTIT                       ".uno:DeleteAnnotation"
#define CMD_SID_DELETEALL_POSTIT                    ".uno:DeleteAllAnnotation"
#define CMD_SID_DELETEALLBYAUTHOR_POSTIT            ".uno:DeleteAllAnnotationByAuthor"
#define CMD_SID_CHARMAP                             ".uno:InsertSymbol"
#define CMD_SID_ATTR_CHAR_POSTURE                   ".uno:Italic"
#define CMD_SID_ATTR_PARA_ADJUST_BLOCK              ".uno:JustifyPara"
#define CMD_SID_LEAVE_GROUP                         ".uno:LeaveGroup"
#define CMD_SID_ATTR_PARA_ADJUST_LEFT               ".uno:LeftPara"
#define CMD_SID_FM_LISTBOX                          ".uno:ListBox"
#define CMD_SID_FM_NUMERICFIELD                     ".uno:NumericField"
#define CMD_SID_OBJECT_ALIGN                        ".uno:ObjectAlign"
#define CMD_SID_FRAME_DOWN                          ".uno:ObjectBackOne"
#define CMD_SID_FRAME_UP                            ".uno:ObjectForwardOne"
#define CMD_SID_FM_OPEN_READONLY                    ".uno:OpenReadOnly"
#define CMD_SID_ATTR_CHAR_CONTOUR                   ".uno:OutlineFont"
#define CMD_SID_PARA_DLG                            ".uno:ParagraphDialog"
#define CMD_SID_FM_PATTERNFIELD                     ".uno:PatternField"
#define CMD_SID_FM_RECORD_SAVE                      ".uno:RecSave"
#define CMD_SID_FM_RECORD_UNDO                      ".uno:RecUndo"
#define CMD_SID_ATTR_PARA_ADJUST_RIGHT              ".uno:RightPara"
#define CMD_SID_FRAME_TO_BOTTOM                     ".uno:SendToBack"
#define CMD_SID_SET_DEFAULT                         ".uno:SetDefault"
#define CMD_SID_ATTR_CHAR_SHADOWED                  ".uno:Shadowed"
#define CMD_SID_SHOW_PROPERTYBROWSER                ".uno:ShowPropBrowser"
#define CMD_SID_FM_SHOW_PROPERTY_BROWSER            ".uno:ShowPropertyBrowser"
#define CMD_SID_ATTR_PARA_LINESPACE_10              ".uno:SpacePara1"
#define CMD_SID_ATTR_PARA_LINESPACE_15              ".uno:SpacePara15"
#define CMD_SID_ATTR_PARA_LINESPACE_20              ".uno:SpacePara2"
#define CMD_SID_ATTR_CHAR_STRIKEOUT                 ".uno:Strikeout"
#define CMD_SID_SET_SUB_SCRIPT                      ".uno:SubScript"
#define CMD_SID_SET_SUPER_SCRIPT                    ".uno:SuperScript"
#define CMD_SID_FM_TAB_DIALOG                       ".uno:TabDialog"
#define CMD_SID_FM_TIMEFIELD                        ".uno:TimeField"
#define CMD_SID_BEZIER_EDIT                         ".uno:ToggleObjectBezierMode"
#define CMD_SID_ATTR_TRANSFORM                      ".uno:TransformDialog"
#define CMD_SID_ATTR_CHAR_UNDERLINE                 ".uno:Underline"
#define CMD_SID_ATTR_CHAR_OVERLINE                  ".uno:Overline"
#define CMD_SID_3D_WIN                              ".uno:Window3D"
#define CMD_SID_OPEN_HYPERLINK                      ".uno:OpenHyperlinkOnCursor"
#define CMD_SID_TABLE_MERGE_CELLS                   ".uno:MergeCells"
#define CMD_SID_TABLE_SPLIT_CELLS                   ".uno:SplitCell"
#define CMD_SID_TABLE_VERT_BOTTOM                   ".uno:CellVertBottom"
#define CMD_SID_TABLE_VERT_CENTER                   ".uno:CellVertCenter"
#define CMD_SID_TABLE_VERT_NONE                     ".uno:CellVertTop"
#define CMD_SID_TABLE_DELETE_ROW                    ".uno:DeleteRows"
#define CMD_SID_TABLE_DELETE_COL                    ".uno:DeleteColumns"
#define CMD_SID_TABLE_SELECT_COL                    ".uno:EntireColumn"
#define CMD_SID_TABLE_SELECT_ROW                    ".uno:EntireRow"
#define CMD_SID_FORMAT_TABLE_DLG                    ".uno:TableDialog"
#define CMD_SID_OPEN_SMARTTAGMENU                   ".uno:OpenSmartTagMenuOnCursor"
#define CMD_SID_TABLE_INSERT_COL_DLG                ".uno:InsertColumnDialog"
#define CMD_SID_TABLE_INSERT_ROW_DLG                ".uno:InsertRowDialog"

#endif

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
