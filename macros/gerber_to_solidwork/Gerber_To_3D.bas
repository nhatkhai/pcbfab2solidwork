Attribute VB_Name = "Gerber_To_3D"
'
'    PCBFab2Solidwork
'    Copyright (C) 2018  NhatKhai L. Nguyen
'
'    Please check LICENSE file for detail.
'
Option Explicit

Public swApp As ISldWorks

Public Const InToMeter = 25.4 / 1000#  ' (mm/in)
Public Const InchToSW As Double = 25.4 / 1000# ' SolidWork use meter unit (m/in)

Public GerbScale As Double      ' =1
Public DrillScale As Double     ' =1

Public POSScale As Double       ' =25.4/1000 Convert to meter - (m/in)
Public AngScale As Double       ' =-1        Convert to angle

Public VRMLScale As Double      ' =25.4/1000 Convert to meter - (m/in)
Public BOMScale As Double       ' =25.4/1000 Convert to meter - (m/in)

Public PCB_Thickness As Double  ' =0.063 (inches)

Public POS_RefColIdx As Integer   '= 0
Public POS_PosXColIdx As Integer  '= 4
Public POS_PosYColIdx As Integer  '= 5
Public POS_RotColIdx As Integer   '= 6
Public POS_SideColIdx As Integer  '= 7
  
Public BOM_RefColIdx        As Integer '= 2
Public BOM_ScaleColIdx      As Integer '= 9
Public BOM_OfsColIdx        As Integer '= 12
Public BOM_RotColIdx        As Integer '= 15
Public BOM_ModleFileColIdx  As Integer '= 18


Const silkTopLayer = "TOP"
Const silkBottomLayer = "BOTTOM"

Const PCBColorR = 0#
Const PCBColorG = 1#
Const PCBColorB = 0#

Const SilkColorR = 0#
Const SilkColorG = 0#
Const SilkColorB = 0#

Const CopperColorR = 1#
Const CopperColorG = 0#
Const CopperColorB = 0#

Const MODEL_SCALEX = 0 ' Scale X
Const MODEL_SCALEY = 1 ' Scale Y
Const MODEL_SCALEZ = 2 ' Scale Z
Const MODEL_OFSX = 3 ' Ofs X
Const MODEL_OFSY = 4 ' Ofs Y
Const MODEL_OFSZ = 5 ' Ofs Z
Const MODEL_ANGX = 6 ' Rot X
Const MODEL_ANGY = 7 ' Rot Y
Const MODEL_ANGZ = 8 ' Rot Z
Const MODEL_FILE = 9 ' 3D Model files (STEP, SLDPRT or VMRL)

'
' Create a new sketch, and bring Solidwork into Edit mode for the sketch
'
' @param Part [in] Solidwork IPardDoc
' @param SketchName [in] Specify the sktech name. This name should be
'        uniqune. Otherwise unexpected behavior from Solidworks may cause
'        unexpected outcome
'
Sub EditNewSketch(Part As IPartDoc _
  , SketchName As String)

  Part.ClearSelection2 True
  Part.SelectionManager.EnableContourSelection = True
  Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0 _
                            , Nothing, 0
  Part.SetUnits swINCHES, swDECIMAL, 8, 3, False
  
  Part.SketchManager.AddToDB = True
  Part.SketchManager.DisplayWhenAdded = False
  Part.SketchManager.Insert3DSketch True
  Part.SketchManager.ActiveSketch.Name = SketchName
End Sub


' This function generate a sketch from a NC Drill file.
'
' Drill format specification can be found at:
' https://web.archive.org/web/20071030075236/http://www.excellon.com/manuals/program.htm
'
' @param absolute_mode [in] set initial assume of Drill file mode. True for
'        absolute, False for relative mode.
'
' @param minHole [in] Only holes with size greater than or equal with
'        this vaule will be generate into sketch
'
' @return A Rect object that containt a minimun board size that would
'        contains all the drill holes
'        
' @note Currently support following commands:
'   INCH 
'   METRIC
'   TZ
'   FMAT,2
'   M48     - Start Drill File
'   M72
'   M71
'   M95, %  - End Header
'   M30
'   G5      - Drill Mode
'   G85     - Slot Mode
'   G90     - Absolute Mode
'   G91     - Incremental Mode
'   T##     - Tool Selection
'   T##C#   - Tool Setting
'   X#Y#    - Coordinate commands
'
Function GenerateSketchFromDrill(Part As IPartDoc _
  , DrillFileName As String _
  , Optional Z As Double = 0 _
  , Optional absolute_mode As Boolean = True _
  , Optional minHole As Double = 0.01) As Rect
 
  Dim inFile  As Integer ' File handle for access Gerber File
  Dim line    As Long    ' Tracking current line in Gerber file
  Dim idx     As Integer ' Tracking current line offset in Gerber file

  Dim ignore  As Boolean ' Use to track and warning user unsupported drill
                         ' commands

  Dim DrillSec As Boolean ' Tracking current line is inside Drill Section 
  
  Dim s As String, ss As String
  Dim x As Double, y As Double
  Dim x1 As Double, y1 As Double

  Dim last_time As Long ' Use to control Solidwork update GUI rate
  
  Dim DrillToSW As Double ' faction_number_unit/SW.unit
  Dim CorrectScale As Double ' Adjust factor for for int -> dbl numbers 
  Dim Leading As Integer
  Dim NumDigit As Integer

  Dim num0 As Integer
  Dim graphic_mode As Integer

  Dim tools(100) As Double ' Tools' radius setting
  Dim drillTool As Integer ' Current Tool number
  Dim r As Double ' Current tool radius

  Dim prevSketch
  
  Dim mySketchMgr As SketchManager
  Set mySketchMgr = Part.SketchManager
  
  ' Initialize variables for tracking minimum size of the board
  Dim MinBrd As Rect
  Set MinBrd = New Rect
  MinBrd.MinX = 1E+20
  MinBrd.MinY = 1E+20
  MinBrd.MaxX = -1E+20
  MinBrd.MaxY = -1E+20
  drillTool = 0
  
  minHole = minHole * InToMeter
  
  ' Read NC Drill file, and sketch drill holes
  line = 0
  x1 = 0
  y1 = 0
  DrillToSW = InchToSW          ' SW/in
  Leading = False
  CorrectScale = DrillScale * 10E-4
  NumDigit     = 6
  graphic_mode = 5 ' Drill Mode
  r = 0
  Set prevSketch = Nothing
  
  inFile = FreeFile
  Open DrillFileName For Input As #inFile

  ' Search the for M48 mark beginning of the Drill header
  DrillSec = False
  Do While Not EOF(inFile)
    Line Input #inFile, s
    s = Trim(s)
    line = line + 1
    If s = "M48" Then 
      DrillSec = True
      Exit Do
    End If
  Loop

  If Not DrillSec Then
    FrmStatus.AppendTODO "Not a valid Drill file"
    FrmStatus.PopTODO
  End If

  ' Process each line in Drill File (after it's header command)
  Do While (Not EOF(inFile)) Or (Not DrillSec)
    Utilities.RelaxForGUI last_time, 0
    Line Input #inFile, s
    s = Trim(s)
    idx = 1
    line = line + 1
    ignore = False
    
    Do
      ss = Utilities.GerberCMD(s, idx)
      Select Case ss
        Case ""
          Exit Do

        Case "M"
          Select Case Utilities.GerberNumber(s, StartIdx:=idx)
            Case 72 ' English Mode (inch)
              DrillToSW = InchToSW           ' SW/in
            Case 71 ' METRIC Mode (mm)
              DrillToSW = InchToSW / 25.4            ' SW/mm
            Case 48 ' Start of the Drill File Header
              ignore = True
            Case 95 ' End of the Drill File Header
              Exit Do
            Case 30 ' End of the Drill File
              DrillSec = False
              Exit Do
            Case Else ' Ignore remain, read next line
              ignore = True
          End Select
        
        Case "INCH" ' English Mode (inch)
          DrillToSW = InchToSW           ' SW/in
          CorrectScale = DrillScale * 10E-4
          NumDigit     = 6
          
        Case "METRIC" ' METRIC Mode (mm)
          DrillToSW = InchToSW / 25.4            ' SW/mm
          CorrectScale = DrillScale * 10E-3
          NumDigit     = 5
          
        Case "TZ"
          Leading = False

        Case "LZ"
          Leading = True
          
        Case "G"
          num0 = Utilities.GerberNumber(s, StartIdx:=idx)
          Select Case num0
            ' Linear, Circular CW, Circular CCW, Variable Dwell, Drill Mode
            Case 1 To 4
              graphic_mode = num0
              ignore = True
            Case 5
              graphic_mode = num0
            Case 85 ' Slot Mode
              If graphic_mode = 5 Then
                prevSketch.Select False
                Part.EditDelete
              End If
              graphic_mode = num0
            Case 90 ' Absolute Mode
              absolute_mode = True
            Case 91 ' Incremental Mode
              absolute_mode = False
            Case Else  ' Ignore remain, read next line
              ignore = True
          End Select
          
        Case "T"
          drillTool = Utilities.GerberNumber(Left(s, idx + 1), StartIdx:=idx)
          r = tools(drillTool)
          
        Case "C"
          r = Utilities.GerberNumber(s, DrillToSW, CorrectScale _
                                    , NumDigit, Leading, idx) / 2#
          tools(drillTool) = r
          
        Case "X", "Y"
          idx = idx - 1
          
          If Not IsNull(Utilities.GerberCMD(s, idx, "X")) Then
            x = Utilities.GerberNumber(s, DrillToSW, CorrectScale _
                                      , NumDigit, Leading, idx)
            If Not absolute_mode Then
              x = x1 + x
            End If
          Else
            x = x1
          End If
          
          If Not IsNull(Utilities.GerberCMD(s, idx, "Y")) Then
            y = Utilities.GerberNumber(s, DrillToSW, CorrectScale _
                                      , NumDigit, Leading, idx)
            If Not absolute_mode Then
              y = y1 + y
            End If
          Else
            y = y1
          End If
          
          If MinBrd.MinX > x - r Then MinBrd.MinX = x - r
          If MinBrd.MaxX < x + r Then MinBrd.MaxX = x + r
          If MinBrd.MinY > y - r Then MinBrd.MinY = y - r
          If MinBrd.MaxY < y + r Then MinBrd.MaxY = y + r
          If r >= minHole Then
            Select Case graphic_mode
              Case 5 ' Drill Mode
                Set prevSketch = mySketchMgr.CreateCircleByRadius(x, y, 0#, r)
              Case 85 ' Slot Mode
                Set prevSketch = mySketchMgr.CreateSketchSlot( _
                    swSketchSlotCreationType_e.swSketchSlotCreationType_line _
                  , swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter _
                  , 2 * r _
                  , x1, y1, 0 _
                  , x, y, 0 _
                  , 0, 0, 0 _
                  , 1, False)
            End Select
          End If
          
          x1 = x
          y1 = y
        
        Case "FMAT"
          Select Case Utilities.GerberNumber(s, StartIdx:=idx) 
            Case 2 ' This script only recognize format version 2
              Exit Do
            Case Else
              ignore = True
          End Select

        Case "%" ' End if Drill Header
          Exit Do

        Case "ICI"
          ignore = True
          
        Case Else  ' Ignore remain, read next line
          If Left(ss, 1) = ";" Then
            Exit Do
          End If
          ignore = True
      End Select

      If ignore Then
        FrmStatus.AppendTODO "Ignore Drill Command " + s + " @ line " + CStr(line)
        FrmStatus.PopTODO
        Exit Do
      End If
    Loop ' Process Dril commands

    If Not DrillSec Then Exit Do
    
  Loop ' Read next Drill command line
  Close #inFile

  Set GenerateSketchFromDrill = MinBrd

End Function ' GenerateSketchFromDrill


' This function generate a sketch from Gerber file.
'
' Geber format specification can be found at:
' https://www.ucamco.com/en/file-formats/gerber/downloads
'
' @param GerberFile [in] Specify full path of a Gerber file
'
' @param Z [in] Specify the SolidWork Z-Plane coordinate for sketch
'             will be generated on. Z default to zero, if omitted.
'
' @NOTE Current support
'   %MO(MM|IN)*            - Units
'   %FS(L|T)(A|?)X\d\d???  - Number format
'   M02 - End of Geber section
'   G04 - Gerber comment
'   G74 - Single Quadrant
'   G75 - Multi Quadrant
'   G90 - Absolute mode
'   G91 - Incremental mode
'   [X#][Y#][I#][J#]D01 - Interpolation
'   [X#][Y#]D02         - Move operation 
'
' @NOTE Need to support
'   G36
'   G37
'   AD
'   D#    Select Aperture Command where #>=10
'   G54D# Select Aperture Command
'   D01 With Aperture
'   D03 With Aperture
'
Sub GenerateSketchFromGerber(Part As IPartDoc _
  , GerberFile As String _
  , Optional Z As Double = 0#)
 
  Dim inFile  As Integer ' File handle for access Gerber File
  Dim line    As Long    ' Tracking current line in Gerber file
  Dim idx     As Integer ' Tracking current line offset in Gerber file
  Dim ignore  As Boolean 
  
  Dim s As String
  Dim x As Double, y As Double
  Dim x1 As Double, y1 As Double
  Dim x2 As Double, y2 As Double
  Dim last_time As Long
  
  Dim GerberToSW As Double ' faction_number_unit/SW.unit
  Dim CorrectScale As Double ' Adjust factor for for int -> dbl numbers 
  Dim NumDigit As Double ' Total Digit format
  
  Dim Leading As Integer, num0 As Integer
  Dim graphic_mode As Integer
  Dim quadrant_mode As Integer
  Dim absolute_mode As Boolean
  Dim dcode As Integer

  Dim prevSketch
  
  Dim mySketchMgr As SketchManager
  Set mySketchMgr = Part.SketchManager
    
  ' Read Silkscreen Gerber file, and sketch silkscreen
  GerberToSW = InchToSW          ' SW/in
  CorrectScale = GerbScale / 1E-4
  NumDigit = 6
  Leading = False
  absolute_mode = True
  graphic_mode = 1
  dcode = 2
  x = 0
  y = 0
  
  line = 0
  inFile = FreeFile
  Open GerberFile For Input As #inFile
  Do While Not EOF(inFile)
    Utilities.RelaxForGUI last_time, 0
    Line Input #inFile, s
    s = Trim(s)
    idx = 1
    line = line + 1
    ignore = False
    
    Select Case Left(s, 1)
      Case "%"
        Select Case Left(s, 3)
          Case "%MO" ' Unit setting
            Select Case Mid(s, 4, 3)
              Case "MM*"  ' Using mm unit
                GerberToSW = InchToSW / 25.4          ' SW/in
              Case "IN*" ' Using Inch unit
                GerberToSW = InchToSW          ' SW/in
              Case Else
                ignore = True
            End Select
            
          Case "%FS"
            If Mid(s, 4, 1) = "L" Then
              Leading = False ' Omit Leading Zero's
            Else
              Leading = True  ' Omit Trailing Zero's
            End If
              
            If Mid(s, 5, 1) = "A" Then
              absolute_mode = True
            Else
              absolute_mode = False
            End If
              
            CorrectScale = GerbScale / (10 ^ CInt(Mid(s, 8, 1)))
            NumDigit = CInt(Mid(s, 7, 1)) + CInt(Mid(s, 8, 1))
        End Select

      Case "M"
        idx = idx + 1
        Select Case Utilities.GerberNumber(s, StartIdx:=idx)
          Case 2 ' M02 - End of Gerber file
            Exit Do
          Case Else
            ignore = True
        End Select
      
      Case "X", "Y", "G"
      
        If Mid(s, idx, 1) = "G" Then
          idx = idx + 1
          num0 = Utilities.GerberNumber(s, StartIdx:=idx)
          Select Case num0
            Case 1, 2, 3 ' G01, G02, G03
              graphic_mode = num0
            Case 4 ' G04 - Comment
              num0 = 4
            Case 74, 75 ' G74, G75
              quadrant_mode = num0
              num0 = 4  ' Ignore the rest
            Case 90 ' Coordinate format to Absolute
              absolute_mode = True
            Case 91 ' Coordinate format to Incremental
              absolute_mode = False
            Case Else
              num0 = 4  ' Ignore the rest
              ignore = True
            End Select
        Else
          num0 = 0
        End If
        
        If num0 <> 4 And Mid(s, idx, 1) <> "*" Then
          'Get X
          If Mid(s, idx, 1) = "X" Then
            idx = idx + 1
            x1 = Utilities.GerberNumber(s, GerberToSW, CorrectScale _
                                       , NumDigit, Leading, idx)
            If Not absolute_mode Then
              x1 = x1 + x
            End If
          Else
            x1 = x
          End If
          
          ' Get Y
          If Mid(s, idx, 1) = "Y" Then
            idx = idx + 1
            y1 = Utilities.GerberNumber(s, GerberToSW, CorrectScale _
                                       , NumDigit, Leading, idx)
            If Not absolute_mode Then
              y1 = y1 + y
            End If
          Else
            y1 = y
          End If
          
          'Get Center X
          If Mid(s, idx, 1) = "I" Then
            idx = idx + 1
            x2 = Utilities.GerberNumber(s, GerberToSW, CorrectScale _
                                       , NumDigit, Leading, idx)
          Else
            x2 = 0
          End If
          
          ' Get Center Y
          If Mid(s, idx, 1) = "J" Then
            idx = idx + 1
            y2 = Utilities.GerberNumber(s, GerberToSW, CorrectScale _
                                       , NumDigit, Leading, idx)
          Else
            y2 = 0
          End If
                      
          ' Get D-Code
          If Mid(s, idx, 1) = "D" Then
            idx = idx + 1
            dcode = Utilities.GerberNumber(s, StartIdx:=idx)
          End If
          
          ' D-Code
          Select case dcode
            Case 1
              Select Case graphic_mode
                Case 1
                  Set prevSketch = mySketchMgr.CreateLine(x, y, Z, x1, y1, Z)
                Case 2, 3 ' Arc Clockwise/CounterClockwise
                  Select Case quadrant_mode
                    Case 75 ' Multi Quadrant
                      Set prevSketch = mySketchMgr.CreateArc(x + x2, y + y2, Z, _
                                        x, y, Z, _
                                        x1, y1, Z, _
                                        (graphic_mode * 2 - 5))
                    Case 74
                      SingleQuadrantArcCenter x, y, x2, y2, x1, y1 _
                                            , (graphic_mode = 2)
                      Set prevSketch = mySketchMgr.CreateArc(x + x2, y + y2, Z, _
                                        x, y, Z, _
                                        x1, y1, Z, _
                                        (graphic_mode * 2 - 5))
                  End Select
              End Select

            Case 2 ' D02 - Move operation

            Case Else
              ignore = True
          End Select
          
          x = x1
          y = y1
        End If

        If num0 <> 4 Then
          If Mid(s, idx, 1) = "*" Then idx = idx+1
          s = Mid(s, idx)
          ignore = Len(s)>0
        End If

      Case Else
        ignore = True
    End Select

    If ignore Then
      FrmStatus.AppendTODO "Ignore Gerber Command " + s + " @ line " + CStr(line)
      FrmStatus.PopTODO
    End If
  Loop ' Process each line in Gerber file
  Close #inFile

  mySketchMgr.DisplayWhenAdded = True
  mySketchMgr.AddToDB = False
End Sub ' GenerateSketchFromGerber


'
' Generate Sketch containt Silk drawing information at specify Z plan
'
' @param Part [in] SolidWork IPartDoc object
' @param SilkFileName [in] Specify full path to drill file
' @param Z [in] Specify were to place the sketch in Z-plane
'
Sub GenerateSilk(Part As IPartDoc _
  , SilkFileName As String _
  , Z As Double _
  , SketchName As String)
 
  FrmStatus.AppendTODO "Create Silkscreen from " + SilkFileName
  EditNewSketch Part, SketchName
  Part.SetLineColor (&HFFFFFF)
  GenerateSketchFromGerber Part, SilkFileName, Z

  ' Close current sketch
  Part.SketchManager.InsertSketch True
  FrmStatus.PopTODO
End Sub ' GenerateSilk


'
' Generated a PCB part from Drill and Board Outline Gerber files.
'
' @param Part [in] Solidwork IPartDoc
' @param DrillFileName [in] Full path to Drill file
' @param GerberFile [in] Full path to the board outline Gerber file
' @param minHole [in] Minimum drill hole size allow to be generate
'
Sub GeneratePCB(Part As IPartDoc _
  , DrillFileName As String _
  , OutLineFileName As String _
  , Optional minHole As Double = 0.01)

  Dim MinBrd As Rect
  Dim brdSp As Double ' Minimum clearance from the drill hole
 
  ' Show User the action plans (first Append last)
  FrmStatus.AppendTODO "Create 3D PCB Part"
  If OutLineFileName <> "" Then
    FrmStatus.AppendTODO "Read Board outline File " + OutLineFileName
  Else
    FrmStatus.AppendTODO "Create estimated board outline"
  End If
  FrmStatus.AppendTODO "Read Drill File " + DrillFileName

  EditNewSketch Part, "Drill"
  Set MinBrd = GenerateSketchFromDrill(Part, DrillFileName)
  Part.SketchManager.InsertSketch True
  FrmStatus.PopTODO
  
  ' Read Board Outline gerber file, and sketch board outline
  EditNewSketch Part, "Board_Outline"
  If OutLineFileName <> "" Then
    GenerateSketchFromGerber Part, OutLineFileName
  Else
    brdSp = 0.1 * InToMeter
    Part.SketchManager.CreateCornerRectangle _
          MinBrd.MinX - brdSp, MinBrd.MinY - brdSp, 0# _
        , MinBrd.MaxX + brdSp, MinBrd.MaxY + brdSp, 0#
  End If
  FrmStatus.PopTODO
  

  Dim featureMgr As IFeatureManager
  Dim myExt as IModelDocExtension
  Dim feature As IFeature

  Set featureMgr = Part.FeatureManager
  Set myExt = Part.Extension

  ' Extruse sketches to generate board from Board Outline
  myExt.SelectByID2 "Board_Outline", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
  myExt.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, True, 16, Nothing, 0
  Set feature = featureMgr.FeatureExtrusion2( True, False, False _
    , swEndCondMidPlane, swEndCondMidPlane _
    , PCB_Thickness * InToMeter, PCB_Thickness * InToMeter _
    , False, False, False, False _
    , 1.74532925199433E-02, 1.74532925199433E-02 _
    , False, False, False, False _
    , False, True, True _
    , swStartOffset, 0, False)
  
  If Not feature Is Nothing Then
    Dim mat
    
    feature.Name = "PCBBoard"
    mat = feature.GetMaterialPropertyValues2(swThisConfiguration, Nothing)
    mat(0) = PCBColorR
    mat(1) = PCBColorG
    mat(2) = PCBColorB
    mat(3) = 0.5 'Ambient
    mat(4) = 1#  'Diffuse
    mat(5) = 0.2 'Specular
    mat(6) = 0.3 'Shininess
    mat(7) = 0#  'Transparency
    mat(8) = 0.2 'Emission
    feature.SetMaterialPropertyValues2 mat, swThisConfiguration, Nothing
  Else
    Part.SketchManager.InsertSketch True
  End If

  ' Cut hold into Board from Drill
  myExt.SelectByID2 "Drill", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
  myExt.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, True, 16, Nothing, 0
  featureMgr.FeatureCut4  True, False, False _
    , swEndCondMidPlane, swEndCondMidPlane _
    , PCB_Thickness * InToMeter, PCB_Thickness * InToMeter _
    , False, False, False, False _
    , 1.74532925199433E-02, 1.74532925199433E-02 _
    , False, False, False, False _
    , False, True, True _
    , True, True, False _
    , swStartOffset, 0, False, False

  FrmStatus.PopTODO
  
  Part.SketchManager.DisplayWhenAdded = True
  Part.SketchManager.AddToDB = False
  Part.Extension.AddComment DrillFileName
End Sub ' GeneratePCB


Sub GenerateVMRL(Part As IPartDoc _
  , FileName As String _
  , Scale_x, Scale_y, Scale_z _
  , Optional genMinMaxBox As Boolean = True _
  , Optional genWireFrame As Boolean = False)

  Dim inFile As Integer
  Dim st As Stack_String
  Dim points As Stack_Dbl
  
  Dim x(3) As Double
  Dim xi As Integer
  Dim prev, s, a, i
  Dim min_x, min_y, min_z, max_x, max_y, max_z
  Dim rmin_x, rmin_y, rmax_x, rmax_y
  Dim r1 As Boolean, r2 As Boolean, r3 As Boolean, r4 As Boolean, r5 As Boolean
  Dim point_count, pp  As Object
  Dim start_idx, end_idx, line, size
  Dim last_time As Long
  
  inFile = FreeFile
  Open FileName For Input As #inFile
  
  Set st = New Stack_String
  Set points = New Stack_Dbl
  
  Dim mySketchMgr As SketchManager
  Set mySketchMgr = Part.SketchManager
  mySketchMgr.AddToDB = True
  mySketchMgr.DisplayWhenAdded = False
  
  Part.SetAddToDB True

  prev = ""
  point_count = 0
  line = 0
  end_idx = 1
  size = 0
  Do
    If end_idx > size Then
      If EOF(inFile) Then Exit Do
      Line Input #inFile, s
      size = Len(s)
      start_idx = 1
      end_idx = 1
    End If
    
    a = Array("", "")
    Do While end_idx <= size
      Select Case Mid(s, end_idx, 1)
        Case "[", "]", ","
          a(0) = Mid(s, start_idx, end_idx - start_idx)
          a(1) = Mid(s, end_idx, 1)
          end_idx = end_idx + 1
          Exit Do
        
        Case " "
          a(0) = Mid(s, start_idx, end_idx - start_idx)
          a(1) = ""
          end_idx = end_idx + 1
          Exit Do
          
        Case vbLf, vbCr
          a(0) = Mid(s, start_idx, end_idx - start_idx)
          a(1) = ""
          line = line + 1
          end_idx = end_idx + 1
          Exit Do
          
        Case Else
          end_idx = end_idx + 1
      End Select
    Loop
    
    Do While end_idx <= Len(s)
      Select Case Mid(s, end_idx, 1)
        Case vbLf, vbCr, " "
          end_idx = end_idx + 1
          
        Case Else
          Exit Do
      End Select
    Loop
    start_idx = end_idx
    
    For i = LBound(a) To UBound(a)
      Utilities.RelaxForGUI last_time, 0
      a(i) = Trim(Replace(a(i), vbLf, " "))
      Select Case a(i)
        Case "["
          If prev = "point" Then
            st.Push "point"
            xi = 0
            min_x = 10000000000000#
            min_y = 10000000000000#
            min_z = 10000000000000#
            max_x = -10000000000000#
            max_y = -10000000000000#
            max_z = -10000000000000#
            If genWireFrame Then
              points.Clear
              mySketchMgr.Insert3DSketch True
            End If
          End If
          
        Case "]"
          If st.Top = "point" Then
            point_count = point_count + 1
            st.Pop
            If genMinMaxBox _
            And ((max_z - min_z) > 0.001 * InToMeter) _
            And ((max_x - min_x) > 0.001 * InToMeter) _
            And ((max_y - min_y) > 0.001 * InToMeter) _
            Then
              Part.ClearSelection2 True
              Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
              mySketchMgr.InsertSketch True
              mySketchMgr.Create3PointCornerRectangle min_x, min_y, 0#, min_x, max_y, 0#, max_x, max_y, 0#
              Set pp = Part.FeatureManager.FeatureExtrusion2( _
                True, False, False, _
                swEndCondMidPlane, swEndCondMidPlane, _
                (max_z - min_z), (max_z - min_z), _
                False, False, False, False, _
                0.01745329251994, 0.01745329251994, _
                False, False, False, False, _
                False, True, True, _
                swStartOffset, (max_z + min_z) / 2, False)
            End If
          End If
          
        Case ""
        
        Case Else
          If st.Top = "point" Then
            If a(i) = "," Then
              If (xi = 3) Then
                x(0) = x(0) * Scale_x * VRMLScale
                x(1) = x(1) * Scale_y * VRMLScale
                x(2) = x(2) * Scale_z * VRMLScale
                If min_x > x(0) Then min_x = x(0)
                If max_x < x(0) Then max_x = x(0)
                If min_y > x(1) Then min_y = x(1)
                If max_y < x(1) Then max_y = x(1)
                If min_z > x(2) Then min_z = x(2)
                If max_z < x(2) Then max_z = x(2)
                If genWireFrame Then
                  points.Push x(0)
                  points.Push x(1)
                  points.Push x(2)
                  If points.Count >= 9 Then
                    Part.CreateLine2 points.GetValue(0), points.GetValue(1), points.GetValue(2), _
                                      points.GetValue(3), points.GetValue(4), points.GetValue(5)
                    Part.CreateLine2 points.GetValue(3), points.GetValue(4), points.GetValue(5), _
                                      points.GetValue(6), points.GetValue(7), points.GetValue(8)
                    Part.CreateLine2 points.GetValue(6), points.GetValue(7), points.GetValue(8), _
                                      points.GetValue(0), points.GetValue(1), points.GetValue(2)
                    points.Pop
                    points.Pop
                    points.Pop
                    points.Pop
                    points.Pop
                    points.Pop
                  End If
                End If
              End If
              xi = 0
            ElseIf IsNumeric(a(i)) Then
              x(xi) = CDbl(a(i))
              xi = xi + 1
            End If
          End If
        
      End Select ' a(i)
      If a(i) <> "" Then prev = a(i)
    Next i
  Loop
  
  Close #inFile

DONE_LABEL:
  mySketchMgr.DisplayWhenAdded = True
  mySketchMgr.AddToDB = False
  Part.Extension.AddComment FileName
End Sub


Sub GeneratePCBAssembly(Part As IAssemblyDoc _
  , boardFilename As String _
  , PosFileName As String _
  , BOMFileName As String _
  , Optional overwriteGeneratedVRML = False _
  , Optional genMinMaxBox As Boolean = True _
  , Optional genWireFrame As Boolean = False _
  , Optional RenameComponents As Boolean = True)
  
  Dim maxCol
  maxCol = POS_SideColIdx
  If maxCol < POS_RotColIdx Then maxCol = POS_RotColIdx
  If maxCol < POS_PosYColIdx Then maxCol = POS_PosYColIdx
  If maxCol < POS_PosXColIdx Then maxCol = POS_PosXColIdx
  If maxCol < POS_RefColIdx Then maxCol = POS_RefColIdx
  
  Dim inFile As Integer
  Dim s As String
  Dim line As Integer
  Dim row, model, models
  
  Dim refs
  
  Dim st As Stack_String          ' Stack keep track of read brd state
  Dim compXforms As Stack_Dbl
  Dim compNames As Stack_String
  Dim compRef As Stack_String
  Dim compVal As Stack_String
  
  Dim MOD_ref, MOD_layer, filter_minSize
  Dim MOD_x, MOD_y, MOD_angle As Double
  Dim OFS_x, OFS_y, OFS_z, Scale_x, Scale_y, Scale_z, Uni_scale As Double
  Dim Rot_x, Rot_y, Rot_z As Double
  Dim Cx, Cy, Cz, Sx, Sy, sz, Ca, Sa As Double
  
  ' These variables used for compute the correct 3D model file
  Dim Na As String, tempStr As String, tempStr2 As String
  Dim extList, ext
  
  Dim mypath As String
  Dim myPart As Object
  Dim start_time As Date
  
  Dim vcomponents As Variant  ' Used for insert parts into the assembly
  Dim tmpComp As Component2   ' Used for rename parts
  Dim seldata As SelectData

  Dim longstatus As Long  ' Dummy number
  Dim last_time As Long   ' Use to relax GUI update
  
  ' Multi-purpose variables
  Dim i As Long
  Dim nErrors As Long, nWarnings As Long
  
  Set st = New Stack_String
  Set compNames = New Stack_String
  Set compXforms = New Stack_Dbl
  Set compRef = New Stack_String
  Set compVal = New Stack_String
  
  FrmStatus.AppendTODO "Rename parts and/or parts' reference"
  FrmStatus.AppendTODO "Insert components"
  FrmStatus.AppendTODO "Read position File" + PosFileName
  
  FrmStatus.AppendTODO "Read 3D BOM File" + BOMFileName
  Set refs = Read3DBOMFile(BOMFileName)
  FrmStatus.PopTODO
  
  filter_minSize = (0.01 * InToMeter) ^ 2
  
  Part.SetAddToDB True
  
  ' Insert PCB
  compNames.Push RemoveFileExt(boardFilename) + ".sldprt"
  
  ' Add a rotational diagonal unit matrix (zero rotation) to the transformation matrix
  compXforms.Push 1#
  compXforms.Push 0#
  compXforms.Push 0#
  compXforms.Push 0#
  compXforms.Push 1#
  compXforms.Push 0#
  compXforms.Push 0#
  compXforms.Push 0#
  compXforms.Push 1#
  
  ' Add a translation vector to the transformation matrix
  compXforms.Push 0#
  compXforms.Push 0#
  compXforms.Push 0#
  
  ' Add a scaling factor to the transform
  compXforms.Push 1#
  
  ' The last three elements of the transformation matrix are unused
  compXforms.Push 0#
  compXforms.Push 0#
  compXforms.Push 0#
  
  ' Add the component to the assembly.
  vcomponents = Part.AddComponents((compNames.GetArray()), (compXforms.GetArray()))
  compNames.Clear
  compXforms.Clear
  Part.ViewZoomtofit2
  
  inFile = FreeFile
  Open PosFileName For Input As #inFile
  
  mypath = FilePath(BOMFileName)
  line = 0
  start_time = Now
  Do While Not EOF(inFile)
    Utilities.RelaxForGUI last_time, 0
    Line Input #inFile, s
    line = line + 1
    If Trim(s) <> "" Then
      row = ReadSpaceSepVecRow(s)
      If UBound(row) >= maxCol Then
        'Extract component position parameters
        MOD_ref = row(POS_RefColIdx)
        If refs.Exists(UCase(MOD_ref)) Then
          For Each model In refs(UCase(MOD_ref)).GetArray()
            MOD_angle = CDbl(row(POS_RotColIdx)) * AngScale
            MOD_x = CDbl(row(POS_PosXColIdx)) * POSScale + 10 * InToMeter
            MOD_y = CDbl(row(POS_PosYColIdx)) * POSScale + 10 * InToMeter
            MOD_layer = UCase(row(POS_SideColIdx))
            If MOD_layer = "PRIMARY SIDE" Then MOD_layer = silkTopLayer
            If MOD_layer = "SECONDARY SIDE" Then MOD_layer = silkBottomLayer
            If MOD_layer = "PRIMARY" Then MOD_layer = silkTopLayer
            If MOD_layer = "SECONDARY" Then MOD_layer = silkBottomLayer
            If MOD_layer = "FRONT" Then MOD_layer = silkTopLayer
            If MOD_layer = "BACK" Then MOD_layer = silkBottomLayer
            If MOD_layer = "BOT" Then MOD_layer = silkBottomLayer
            
            ' Extract 3D model parameters
            tempStr = RemoveFileExt(model(MODEL_FILE))
            ext = GetFileExt(model(MODEL_FILE))
            Scale_x = model(MODEL_SCALEX)
            Scale_y = model(MODEL_SCALEY)
            Scale_z = model(MODEL_SCALEZ)
            OFS_x = model(MODEL_OFSX) * BOMScale
            OFS_y = model(MODEL_OFSY) * BOMScale
            OFS_z = model(MODEL_OFSZ) * BOMScale
            Rot_x = model(MODEL_ANGX)
            Rot_y = model(MODEL_ANGY)
            Rot_z = model(MODEL_ANGZ)
            
            If MOD_layer <> "" And tempStr <> "" Then
              ' Check for uniform scaleing
              If Abs(Scale_x - Scale_y) < 0.0002 _
              And Abs(Scale_x - Scale_z) < 0.0002 Then
                tempStr2 = ""
                Uni_scale = Scale_x
                Scale_x = 1
                Scale_y = 1
                Scale_z = 1
              Else
                tempStr2 = "_" + Str(Scale_x) + _
                          "_" + Str(Scale_y) + _
                          "_" + Str(Scale_z)
                Uni_scale = 1
              End If
        
              ' Search for model files
              If genMinMaxBox Then
                  extList = Array(ext, ".wrl", ".step", ".stp", "")
              Else
                  extList = Array(ext, ".step", ".stp", ".wrl", "")
              End If
              
              For Each ext In extList
                Na = FindFile(mypath, tempStr, ext)
                If Na <> "" Then
                  Exit For
                End If
              Next
              
              ' Check if base solidwork model already exist
              If tempStr2 <> "" Then
                extList = FindFile(mypath, tempStr, ".sldprt")
                If extList <> "" Then
                  If FileDateTime(Na + ext) <= FileDateTime(extList + ".sldprt") Then
                    ext = ".sldprt"
                  End If
                End If
              End If
              
              ' Check if Solidwork model already exist, or need to regenerate it from source
              extList = FindFile(mypath, tempStr + tempStr2, ".sldprt")
              If extList <> "" Then
                If overwriteGeneratedVRML Then
                  ext = "USE_SLDPRT"
                ElseIf start_time < FileDateTime(extList + ".sldprt") Then
                  ext = "USE_SLDPRT"
                ElseIf ext <> "" _
                  And FileDateTime(Na + ext) <= FileDateTime(extList + ".sldprt") Then
                  ext = "USE_SLDPRT"
                End If
              End If
              
              Select Case ext
                Case ""
                  Na = ""
                  
                Case "USE_SLDPRT"
                  Na = extList + ".sldprt"
                  
                  Case ".sldprt"
                    FrmStatus.AppendTODO MOD_ref + " use 3D model from " + Na + ext
                    
                    Set myPart = swApp.OpenDoc6(Na + ext, swDocPART, _
                      swOpenDocOptions_ReadOnly Or swOpenDocOptions_Silent, _
                      "", nErrors, nWarnings)
                    myPart.Visible = PCBCfgForm.PartVisible
                    myPart.FeatureManager.InsertScale 0, False, Scale_x, Scale_y, Scale_z
                    myPart.ViewZoomtofit2
                    Na = Na + tempStr2 + ".sldprt"
                    myPart.SaveAs Na
                    swApp.CloseDoc GetFileName(myPart.GetPathName())
                    Set myPart = Nothing
                    
                    FrmStatus.PopTODO
                  
                  
                  Case ".wrl"
                    Na = Na
                    FrmStatus.AppendTODO MOD_ref + " use 3D model from " + Na + ext
                    
                    Set myPart = swApp.NewPart
                    myPart.Visible = PCBCfgForm.PartVisible
                    GenerateVMRL myPart, Na + ext, Scale_x, Scale_y, Scale_z, True, genWireFrame
                    myPart.ViewZoomtofit2
                    Na = Na + tempStr2 + ".sldprt"
                    myPart.SaveAs Na
                    swApp.CloseDoc GetFileName(myPart.GetPathName())
                    Set myPart = Nothing
                    
                    FrmStatus.PopTODO
                  
                  Case Else
                    ' Check if STEP file exist
                    Na = Na
                    FrmStatus.AppendTODO MOD_ref + " use 3D model from " + Na + ext
                    
                    Dim importData As SldWorks.ImportStepData
                    Set importData = swApp.GetImportFileData(Na + ext)
                    importData.MapConfigurationData = True
                    Set myPart = swApp.LoadFile4(Na + ext, "r", importData, longstatus)
                    If (Scale_x = Scale_y) And (Scale_x = Scale_z) Then
                      If Abs(Scale_x - 1) > 0.001 Then
                        myPart.FeatureManager.InsertScale swScaleAboutOrigin, True, Scale_x, Scale_y, Scale_z
                      End If
                    Else
                      myPart.FeatureManager.InsertScale swScaleAboutOrigin, False, Scale_x, Scale_y, Scale_z
                    End If
                    myPart.Visible = PCBCfgForm.PartVisible
                    myPart.ViewZoomtofit2
                    Na = Na + tempStr2 + ".sldprt"
                    myPart.SaveAs Na
                    swApp.CloseDoc GetFileName(myPart.GetPathName())
                    Set myPart = Nothing
                    
                    FrmStatus.PopTODO
              End Select
                    
              If Na <> "" Then
                compNames.Push Na
                compRef.Push MOD_ref
                compVal.Push RemoveFileExt(GetFileName(Na))
            
                ' Define the transformation matrix. See the IMathTransform API documentation.
                ' Add a rotational diagonal unit matrix (zero rotation) to the transformation matrix
                'mod_angle = 180 - mod_angle
                If MOD_layer = silkBottomLayer Then
                  Rot_x = Rot_x + 180
                  Rot_y = -Rot_y
                  Rot_z = -MOD_angle - Rot_z
                  MOD_angle = 180 + MOD_angle
                  OFS_x = -OFS_x
                Else
                  Rot_z = -MOD_angle + Rot_z
                  MOD_angle = MOD_angle
                End If
                              
                Rot_x = Rot_x / 180 * PI
                Rot_y = Rot_y / 180 * PI
                Rot_z = Rot_z / 180 * PI
                MOD_angle = MOD_angle / 180 * PI
                Cx = Cos(Rot_x)
                Sx = Sin(Rot_x)
                Cy = Cos(Rot_y)
                Sy = Sin(Rot_y)
                Cz = Cos(Rot_z)
                sz = Sin(Rot_z)
                Ca = Cos(MOD_angle)
                Sa = Sin(MOD_angle)
        
                compXforms.Push Cy * Cz
                compXforms.Push -Cy * sz
                compXforms.Push Sy
                compXforms.Push Sx * Sy * Cz + Cx * sz
                compXforms.Push Cx * Cz - Sx * Sy * sz
                compXforms.Push -Sx * Cy
                compXforms.Push Sx * sz - Cx * Sy * Cz
                compXforms.Push Cx * Sy * sz + Sx * Cz
                compXforms.Push Cx * Cy
                
                ' Add a translation vector to the transformation matrix
                compXforms.Push OFS_x * Ca - OFS_y * Sa + MOD_x
                compXforms.Push OFS_y * Ca + OFS_x * Sa + MOD_y
                If MOD_layer = silkTopLayer Then
                  compXforms.Push OFS_z + PCB_Thickness / 2 * InToMeter
                Else
                  compXforms.Push -(OFS_z + PCB_Thickness / 2 * InToMeter)
                End If
                
                ' Add a scaling factor to the transform
                compXforms.Push Uni_scale
                
                ' The last three elements of the transformation matrix are unused
                compXforms.Push 0#
                compXforms.Push 0#
                compXforms.Push 0#
              
              End If
            End If
          Next model
        End If
      End If
    End If
  Loop
  FrmStatus.PopTODO
  
  Close #inFile

  ' Add the component to the assembly.
  FrmStatus.setMaxValue compNames.Count
  FrmStatus.setRemaindValue compNames.Count
  If Not compNames.IsEmpty() Then
    vcomponents = Part.AddComponents((compNames.GetArray()), (compXforms.GetArray()))
    FrmStatus.PopTODO
  
    FrmStatus.setMaxValue compNames.Count
    For i = LBound(vcomponents) To UBound(vcomponents)
      FrmStatus.setCurrentValue i
      Utilities.RelaxForGUI last_time, 0
      Set tmpComp = vcomponents(i)
      tempStr = Replace(compRef.GetValue(i), "@", "at")
      tmpComp.ComponentReference = tempStr
      If RenameComponents Then
        tempStr2 = tempStr + "_" + Replace(compVal.GetValue(i), "@", "at")
        tempStr2 = Replace(tempStr2, "/", ",")
        'tmpComp.Select2 False, 0
        tmpComp.Select4 False, seldata, False
        tmpComp.Name2 = tempStr2
      End If
    Next i
    FrmStatus.PopTODO
  Else
    FrmStatus.PopTODO
    FrmStatus.PopTODO
  End If
  
  compNames.Clear
  compXforms.Clear
  compRef.Clear
  compVal.Clear
End Sub


Sub main()
  Set swApp = Application.SldWorks

  ' Intialize internal and GUI parameters if not initialized
  If PCBCfgForm.DrillFileName = "" Then
    PCBCfgForm.DrillFileName = ""
    PCBCfgForm.OutLineFileName = ""
    PCBCfgForm.TopSilkFileName = ""
    PCBCfgForm.BotSilkFileName = ""
    PCBCfgForm.PosFileName = ""
    PCBCfgForm.BOMFileName = ""
    
    DrillScale = 1#  ' unit/in
    GerbScale = 1#   ' unit/in
    
    POSScale = InchToSW / 1000#     ' unit/in
    AngScale = 1                    ' unit/degree
    POS_RefColIdx = 0
    POS_PosXColIdx = 4
    POS_PosYColIdx = 5
    POS_RotColIdx = 6
    POS_SideColIdx = 7
    
    VRMLScale = InchToSW  ' unit/in
    PCB_Thickness = 0.063 ' in
    
    BOMScale = InchToSW   ' unit/in
    BOM_RefColIdx = 0
    BOM_ScaleColIdx = 2
    BOM_OfsColIdx = 5
    BOM_RotColIdx = 8
    BOM_ModleFileColIdx = 11
    
    PCBCfgForm.txtDrillScale = CStr(1 / DrillScale)
    PCBCfgForm.txtGerbScale = CStr(1 / GerbScale)
    PCBCfgForm.txtPosScale = CStr(InchToSW / POSScale)
    PCBCfgForm.txtPosAngleScale = CStr(AngScale)
    PCBCfgForm.txtWRLScale = CStr(InchToSW / VRMLScale)
    PCBCfgForm.txtPCBThickness = CStr(PCB_Thickness * 1000)
  End If
  PCBCfgForm.Show
End Sub
