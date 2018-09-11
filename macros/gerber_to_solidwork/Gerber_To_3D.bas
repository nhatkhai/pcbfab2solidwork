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

Public GerbScale As Double      ' =25.4      Convert to meter - (m/in)
Public DrillScale As Double     ' =25.4/1000 Convert to meter - (m/in)

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

Sub GenerateSilk(Part As IPartDoc, SilkFileName As String _
  , Z As Double _
  , Optional SketchName As String = "")
 
  Dim inFile As Integer
  Dim line As Long
  Dim idx As Integer
  
  Dim s As String, a() As String, ss As String
  Dim x As Double, y As Double
  Dim x1 As Double, y1 As Double
  Dim x2 As Double, y2 As Double
  Dim last_time As Long
  Dim mySketch As SketchManager
  Dim myfeature As FeatureManager
  Dim mat
  
  Dim DrillToSW As Double ' faction_number_unit/SW.unit
  Dim GerberNumberToSW As Double ' gerber_number_unit/SW.unit
  Dim Leading As Integer, num0 As Integer
  Dim graphic_mode As Integer
  Dim quadrant_mode As Integer
  Dim absolute_mode As Boolean
  Dim dcode As Integer
  Dim radius(100) As Double
  Dim r As Double
  Dim prevSketch
  
  Part.ClearSelection2 True
  Part.SelectionManager.EnableContourSelection = True
  Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
  Part.SetUnits swINCHES, swDECIMAL, 8, 3, False
  
  Set myfeature = Part.FeatureManager
  Set mySketch = Part.SketchManager
  mySketch.AddToDB = True
  mySketch.DisplayWhenAdded = False
  mySketch.Insert3DSketch True
  If SketchName = "" Then
    mySketch.ActiveSketch.Name = "Silk_" + CStr(Z / InchToSW * 1000) + "mil"
  Else
    mySketch.ActiveSketch.Name = SketchName
  End If
  Part.SetLineColor (16777215)
  
  FrmStatus.AppendTODO "Create Silkscree from " + SilkFileName
 
  ' Read Silkscree gerber file, and sketch silkscreen
  DrillToSW = InchToSW          ' SW/in
  GerberNumberToSW = GerbScale  ' SW/unit
  Leading = False
  absolute_mode = True
  graphic_mode = 1
  dcode = 2
  x = 0
  y = 0
  
  line = 0
  inFile = FreeFile
  Open SilkFileName For Input As #inFile
  Do While Not EOF(inFile)
    RelaxForGUI last_time, 0
    Line Input #inFile, s
    s = Trim(s)
    idx = 1
    line = line + 1
    
    Select Case Left(s, 1)
      Case "%"
        Select Case Left(s, 3)
          Case "%MO" ' Unit setting
            Select Case Mid(s, 4, 3)
              Case "MM*"  ' Using mm unit
                DrillToSW = InchToSW / 25.4          ' SW/in
                GerberNumberToSW = GerbScale / 25.4  ' SW/unit
              Case "IN*" ' Using Inch unit
                DrillToSW = InchToSW          ' SW/in
                GerberNumberToSW = GerbScale  ' SW/unit
              End Select
            
          Case "%FS"
            If Mid(s, 4, 1) = "L" Then
              Leading = True
            Else
              Leading = False
            End If
              
            If Mid(s, 5, 1) = "A" Then
              absolute_mode = True
            Else
              absolute_mode = False
            End If
              
            GerberNumberToSW = DrillToSW / (10 ^ CInt(Mid(s, 8, 1)))
        End Select
      
      Case "X", "Y", "G"
      
        If Mid(s, idx, 1) = "G" Then
          idx = idx + 1
          num0 = Utilities.GerberNumber(s, 1, 1, , True, idx)
          If num0 <= 3 Then
            ' G01, G02, G03
            graphic_mode = num0
          ElseIf num0 < 74 Then
            num0 = 4 ' Ignore the rest
          ElseIf num0 <= 75 Then
            ' G74, G75
            quadrant_mode = num0
            num0 = 4  ' Ignore the rest
          Else
            num0 = 4  ' Ignore the rest
          End If
        Else
          num0 = 0
        End If
        
        If num0 <> 4 And Mid(s, idx, 1) <> "*" Then
          'Get X
          If Mid(s, idx, 1) = "X" Then
            idx = idx + 1
            x1 = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , Leading, idx)
            If Not absolute_mode Then
              x1 = x1 + x
            End If
          Else
            x1 = x
          End If
          
          ' Get Y
          If Mid(s, idx, 1) = "Y" Then
            idx = idx + 1
            y1 = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , Leading, idx)
            If Not absolute_mode Then
              y1 = y1 + y
            End If
          Else
            y1 = y
          End If
          
          'Get Center X
          If Mid(s, idx, 1) = "I" Then
            idx = idx + 1
            x2 = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , Leading, idx)
          Else
            x2 = 0
          End If
          
          ' Get Center Y
          If Mid(s, idx, 1) = "J" Then
            idx = idx + 1
            y2 = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , Leading, idx)
          Else
            y2 = 0
          End If
                      
          ' Get D-Code
          If Mid(s, idx, 1) = "D" Then
            idx = idx + 1
            dcode = Utilities.GerberNumber(s, 1, 1, , True, idx)
          End If
          
          ' D-Code
          If dcode = 1 Then
            Select Case graphic_mode
              Case 1
                Set prevSketch = mySketch.CreateLine(x, y, Z, x1, y1, Z)
              Case 2, 3 ' Arc Clockwise/CounterClockwise
                Select Case quadrant_mode
                  Case 75 ' Multi Quadrant
                    Set prevSketch = mySketch.CreateArc(x + x2, y + y2, Z, _
                                       x, y, Z, _
                                       x1, y1, Z, _
                                       (graphic_mode * 2 - 5))
                  Case 74
                    SingleQuadrantArcCenter x, y, x2, y2, x1, y1, (graphic_mode = 2)
                    Set prevSketch = mySketch.CreateArc(x + x2, y + y2, Z, _
                                       x, y, Z, _
                                       x1, y1, Z, _
                                       (graphic_mode * 2 - 5))
                End Select
            End Select
          End If
          
          x = x1
          y = y1
        End If
    End Select
  Loop
  Close #inFile
  
  FrmStatus.PopTODO
  mySketch.InsertSketch True
  mySketch.DisplayWhenAdded = True
  mySketch.AddToDB = False
  Part.Extension.AddComment SilkFileName
End Sub



Sub GeneratePCB(Part As IPartDoc, FileName As String, OutLineFileName As String, _
 Optional absolute_mode As Boolean = True, _
 Optional minHole As Double = 0.01)
 
  Dim inFile As Integer
  Dim line As Long
  Dim idx As Integer
  
  Dim s As String, a() As String, ss As String
  Dim x As Double, y As Double
  Dim x1 As Double, y1 As Double
  Dim x2 As Double, y2 As Double
  Dim last_time As Long
  Dim mySketch As SketchManager
  Dim myfeature As FeatureManager
  Dim mat
  
  Dim boardMinX As Double
  Dim boardMinY As Double
  Dim boardMaxX As Double
  Dim boardMaxY As Double
  Dim brdSp As Double
  
  Dim DrillToSW As Double ' faction_number_unit/SW.unit
  Dim GerberNumberToSW As Double ' gerber_number_unit/SW.unit
  Dim Leading As Integer, num0 As Integer
  Dim graphic_mode As Integer
  Dim quadrant_mode As Integer
  Dim dcode As Integer
  Dim drillTool As Integer
  Dim radius(100) As Double
  Dim r As Double
  Dim prevSketch
  
  Part.ClearSelection2 True
  Part.SelectionManager.EnableContourSelection = True
  Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
  Part.SetUnits swINCHES, swDECIMAL, 8, 3, False
  
  Set myfeature = Part.FeatureManager
  Set mySketch = Part.SketchManager
  mySketch.AddToDB = True
  mySketch.DisplayWhenAdded = False
  mySketch.InsertSketch True
  mySketch.ActiveSketch.Name = "Sketch_PCBBoard"
  
  boardMinX = 10000000000000#
  boardMinY = 10000000000000#
  boardMaxX = -10000000000000#
  boardMaxY = -10000000000000#
  drillTool = 0
  
  minHole = minHole * InToMeter
  brdSp = 0.1 * InToMeter
  
  FrmStatus.AppendTODO "Create 3D PCB Part"
  If OutLineFileName <> "" Then
    FrmStatus.AppendTODO "Read Board outline File " + OutLineFileName
  Else
    FrmStatus.AppendTODO "Create estimated board outline"
  End If
  FrmStatus.AppendTODO "Read Drill File " + FileName
  
  ' Read NC Drill file, and sketch drill holes
  line = 0
  x1 = 0
  y1 = 0
  DrillToSW = InchToSW          ' SW/in
  GerberNumberToSW = DrillScale ' SW/unit
  Leading = False
  graphic_mode = 5 ' Drill Mode
  r = 0
  Set prevSketch = Nothing
  
  inFile = FreeFile
  Open FileName For Input As #inFile
  Do While Not EOF(inFile)
    RelaxForGUI last_time, 0
    Line Input #inFile, s
    s = Trim(s)
    idx = 1
    line = line + 1
    
    Do
      ss = Utilities.GerberCMD(s, idx)
      Select Case ss
        Case "M"
          Select Case Utilities.GerberNumber(s, 1, 1, , True, idx)
            Case 72 ' English Mode (inch)
              DrillToSW = InchToSW           ' SW/in
              GerberNumberToSW = DrillScale  ' SW/unit
            Case 71 ' METRIC Mode (mm)
              DrillToSW = InchToSW / 25.4            ' SW/mm
              GerberNumberToSW = DrillScale / 25.4 ' SW/unit
            Case Else ' Ignore remain, read next line
              Exit Do
          End Select
        
        Case "METRIC" ' METRIC Mode (mm)
          DrillToSW = InchToSW / 25.4            ' SW/mm
          GerberNumberToSW = DrillScale / 25.4 ' SW/unit
          
        Case "INCH" ' English Mode (inch)
          DrillToSW = InchToSW           ' SW/in
          GerberNumberToSW = DrillScale  ' SW/unit
          
        Case "TZ"
          Leading = True
          
        Case "G"
          num0 = Utilities.GerberNumber(s, 1, 1, , True, idx)
          Select Case num0
            Case 1 To 5 ' Linear, Circular CW, Circular CCW, Variable Dwell, Drill Mode
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
            Case 92, 93 ' Ingore Set Zero Command
              Exit Do
            Case Else  ' Ignore remain, read next line
              GoTo IgnoreDrill
          End Select
          
        Case "T"
          drillTool = Utilities.GerberNumber(Left(s, idx + 1), 1, 1, , True, idx)
          Utilities.GerberNumber Left(s, idx + 1), 1, 1, , True, idx
          r = radius(drillTool)
          
        Case "C"
          r = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , True, idx) / 2#
          radius(drillTool) = r
          
        Case "X", "Y"
          idx = idx - 1
          
          If Not IsNull(Utilities.GerberCMD(s, idx, "X")) Then
            x = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , Leading, idx)
            If Not absolute_mode Then
              x = x1 + x
            End If
          Else
            x = x1
          End If
          
          If Not IsNull(Utilities.GerberCMD(s, idx, "Y")) Then
            y = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , Leading, idx)
            If Not absolute_mode Then
              y = y1 + y
            End If
          Else
            y = y1
          End If
          
          If boardMinX > x - brdSp - r Then boardMinX = x - brdSp - r
          If boardMaxX < x + brdSp + r Then boardMaxX = x + brdSp + r
          If boardMinY > y - brdSp - r Then boardMinY = y - brdSp - r
          If boardMaxY < y + brdSp + r Then boardMaxY = y + brdSp + r
          If r >= minHole Then
            Select Case graphic_mode
              Case 5 ' Drill Mode
                Set prevSketch = mySketch.CreateCircleByRadius(x, y, 0#, r)
              Case 85 ' Slot Mode
                Set prevSketch = mySketch.CreateSketchSlot( _
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
        
        Case "FMAT", "%", "", "ICI"
          Exit Do
          
        Case Else  ' Ignore remain, read next line
          If Left(ss, 1) = ";" Then
            Exit Do
          End If
          GoTo IgnoreDrill
      End Select
    Loop
    GoTo DrillDone

IgnoreDrill:
        FrmStatus.AppendTODO "Ignore Drill Command " + s + " @ line " + CStr(line)
        FrmStatus.PopTODO
DrillDone:
        On Error GoTo 0
    
  Loop
  Close #inFile
  FrmStatus.PopTODO
  
  ' Read Board Outline gerber file, and sketch board outline
  If OutLineFileName <> "" Then
    DrillToSW = InchToSW          ' SW/in
    GerberNumberToSW = GerbScale  ' SW/unit
    Leading = False
    absolute_mode = True
    graphic_mode = 1
    dcode = 2
    
    line = 0
    inFile = FreeFile
    Open OutLineFileName For Input As #inFile
    x = 0
    y = 0
    Do While Not EOF(inFile)
      RelaxForGUI last_time, 0
      Line Input #inFile, s
      s = Trim(s)
      idx = 1
      line = line + 1
      
      Select Case Left(s, 1)
        Case "%"
          Select Case Left(s, 3)
            Case "%MO" ' Unit setting
              Select Case Mid(s, 4, 3)
                Case "MM*"  ' Using mm unit
                  DrillToSW = InchToSW / 25.4          ' SW/in
                  GerberNumberToSW = GerbScale / 25.4  ' SW/unit
                Case "IN*" ' Using Inch unit
                  DrillToSW = InchToSW          ' SW/in
                  GerberNumberToSW = GerbScale  ' SW/unit
                End Select
              
            Case "%FS"
              If Mid(s, 4, 1) = "L" Then
                Leading = True
              Else
                Leading = False
              End If
                
              If Mid(s, 5, 1) = "A" Then
                absolute_mode = True
              Else
                absolute_mode = False
              End If
                
              GerberNumberToSW = DrillToSW / (10 ^ CInt(Mid(s, 8, 1)))
          End Select
        
        Case "X", "Y", "G"
        
          If Mid(s, idx, 1) = "G" Then
            idx = idx + 1
            num0 = Utilities.GerberNumber(s, 1, 1, , True, idx)
            If num0 <= 3 Then
              ' G01, G02, G03
              graphic_mode = num0
            ElseIf num0 < 74 Then
              num0 = 4 ' Ignore the rest
            ElseIf num0 <= 75 Then
              ' G74, G75
              quadrant_mode = num0
              num0 = 4  ' Ignore the rest
            Else
              num0 = 4  ' Ignore the rest
            End If
          Else
            num0 = 0
          End If
          
          If num0 <> 4 And Mid(s, idx, 1) <> "*" Then
            'Get X
            If Mid(s, idx, 1) = "X" Then
              idx = idx + 1
              x1 = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , Leading, idx)
              If Not absolute_mode Then
                x1 = x1 + x
              End If
            Else
              x1 = x
            End If
            
            ' Get Y
            If Mid(s, idx, 1) = "Y" Then
              idx = idx + 1
              y1 = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , Leading, idx)
              If Not absolute_mode Then
                y1 = y1 + y
              End If
            Else
              y1 = y
            End If
            
            'Get Center X
            If Mid(s, idx, 1) = "I" Then
              idx = idx + 1
              x2 = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , Leading, idx)
            Else
              x2 = 0
            End If
            
            ' Get Center Y
            If Mid(s, idx, 1) = "J" Then
              idx = idx + 1
              y2 = Utilities.GerberNumber(s, DrillToSW, GerberNumberToSW, , Leading, idx)
            Else
              y2 = 0
            End If
                        
            ' Get D-Code
            If Mid(s, idx, 1) = "D" Then
              idx = idx + 1
              dcode = Utilities.GerberNumber(s, 1, 1, , True, idx)
            End If
            
            ' D-Code
            If dcode = 1 Then
              Select Case graphic_mode
                Case 1
                  mySketch.CreateLine x, y, 0#, x1, y1, 0#
                Case 2, 3 ' Arc Clockwise/CounterClockwise
                  Select Case quadrant_mode
                    Case 75 ' Multi Quadrant
                      mySketch.CreateArc x + x2, y + y2, 0#, _
                                         x, y, 0#, _
                                         x1, y1, 0#, _
                                         (graphic_mode * 2 - 5)
                    Case 74
                      SingleQuadrantArcCenter x, y, x2, y2, x1, y1, (graphic_mode = 2)
                      mySketch.CreateArc x + x2, y + y2, 0#, _
                                         x, y, 0#, _
                                         x1, y1, 0#, _
                                         (graphic_mode * 2 - 5)
                  End Select
              End Select
            End If
            
            x = x1
            y = y1
          End If
      End Select
      
    Loop
    Close #inFile
  Else
    mySketch.CreateCornerRectangle boardMinX, boardMinY, 0#, boardMaxX, boardMaxY, 0#
  End If
  FrmStatus.PopTODO
  
  ' Extruse sketches to generate board part
  Dim feature As IFeature, tf
  Set feature = myfeature.FeatureExtrusion2( _
    True, False, False, _
    swEndCondMidPlane, swEndCondMidPlane, _
    PCB_Thickness * InToMeter, PCB_Thickness * InToMeter, _
    False, False, False, False, _
    0.01745329251994, 0.01745329251994, _
    False, False, False, False, _
    False, True, True, _
    swStartOffset, 0, False)
  
  If Not feature Is Nothing Then
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
    mySketch.InsertSketch True
  End If
  FrmStatus.PopTODO
  
  mySketch.DisplayWhenAdded = True
  mySketch.AddToDB = False
  Part.Extension.AddComment FileName
End Sub

Sub SketchRecSilk(mySketch As SketchManager, silks As Stack_Dbl, z1 As Double)
  Dim x1 As Double, y1 As Double
  Dim x2 As Double, y2 As Double
  Dim x3 As Double, y3 As Double
  Dim x4 As Double, y4 As Double
  Dim r1 As Double
  Dim MOD_x As Double, MOD_y As Double
  Dim angle As Double
  Dim last_time As Long

  FrmStatus.setMaxValue silks.Count
  
  Do While Not silks.IsEmpty()
    FrmStatus.setRemaindValue silks.Count
    RelaxForGUI last_time, 0
    
    y1 = silks.Pop()
    x1 = silks.Pop()
    y2 = silks.Pop()
    x2 = silks.Pop()
    r1 = silks.Pop()
    angle = silks.Pop()
    MOD_y = silks.Pop()
    MOD_x = silks.Pop()
    
    If (r1 < 0.008 * InToMeter) Then
      r1 = 0.008 * InToMeter
    End If
    x3 = x1 + r1
    y3 = y1 + r1
    x4 = x2 - r1
    y4 = y2 - r1
    If ((x4 - x3 < 0.005 * InToMeter) Or (y4 - y3 < 0.005 * InToMeter)) Then
      r1 = 0
    End If
      
    Rotate2D x1, y1, angle, MOD_x, MOD_y
    Rotate2D x2, y2, angle, MOD_x, MOD_y
    Rotate2D x3, y3, angle, MOD_x, MOD_y
    Rotate2D x4, y4, angle, MOD_x, MOD_y
    
    mySketch.Create3PointCornerRectangle x1, y1, z1, x1, y2, z1, x2, y2, z1
    If (r1 > 0) Then
      mySketch.Create3PointCornerRectangle x3, y3, z1, x3, y4, z1, x4, y4, z1
    End If
  Loop
End Sub

Sub SketchDSSilk(mySketch As SketchManager, silks As Stack_Dbl, z1 As Double, _
  Optional myfeature As IFeatureManager = Nothing, _
  Optional Name As String = "")
  
  Dim x1 As Double, y1 As Double
  Dim x2 As Double, y2 As Double
  Dim r1 As Double
  Dim MOD_x As Double, MOD_y As Double
  Dim MOD_angle As Double, angle As Double
  Dim last_time As Long
  Dim feature As IFeature
  Dim mat

  FrmStatus.setMaxValue silks.Count
  
  Do While Not silks.IsEmpty()
    FrmStatus.setRemaindValue silks.Count
    RelaxForGUI last_time, 0
    
    r1 = silks.Pop()
    x1 = silks.Pop()
    y1 = silks.Pop()
    x2 = silks.Pop()
    y2 = silks.Pop()
    angle = silks.Pop()
    MOD_x = silks.Pop()
    MOD_y = silks.Pop()
    
    If (Abs(x1 - x2) > r1 Or Abs(y1 - y2) > r1) And (r1 >= 0.005 * InToMeter) Then
      Rotate2D x1, y1, angle, MOD_x, MOD_y
      Rotate2D x2, y2, angle, MOD_x, MOD_y
      
      mySketch.CreateSketchSlot _
        swSketchSlotCreationType_e.swSketchSlotCreationType_line, _
        swSketchSlotLengthType_e.swSketchSlotLengthType_CenterCenter, _
        r1, _
        x1, y1, z1, _
        x2, y2, z1, _
        0, 0, 0, _
        1, False
      If Not myfeature Is Nothing Then
        mySketch.ActiveSketch.Name = "Sketch_" + Name + "_" + Str(silks.Count / 8)
        Set feature = myfeature.FeatureExtrusion2( _
          True, False, False, _
          swEndCondMidPlane, swEndCondMidPlane, _
          0.001 * InToMeter, 0.001 * InToMeter, _
          False, False, False, False, _
          0.01745329251994, 0.01745329251994, _
          False, False, False, False, _
          False, True, True, _
          swStartOffset, z1, False)
        If Not feature Is Nothing Then
          feature.Name = Name + "_" + Str(silks.Count / 8)
          mat = feature.GetMaterialPropertyValues2(swThisConfiguration, Nothing)
          mat(0) = SilkColorR
          mat(1) = SilkColorG
          mat(2) = SilkColorB
          mat(3) = 0.5 'Ambient
          mat(4) = 1#  'Diffuse
          mat(5) = 0.2 'Specular
          mat(6) = 0.3 'Shininess
          mat(7) = 0#  'Transparency
          mat(8) = 0.6 'Emission
          feature.SetMaterialPropertyValues2 mat, swThisConfiguration, Nothing
        Else
          mySketch.InsertSketch True
        End If
      End If
    End If
  Loop
End Sub

Sub SketchDCSilk(mySketch As SketchManager, silkDC As Stack_Dbl, z1 As Double)
  Dim x1 As Double, y1 As Double
  Dim x2 As Double, y2 As Double
  Dim r1 As Double
  Dim MOD_x As Double, MOD_y As Double
  Dim MOD_angle As Double, angle As Double
  Dim last_time As Long

  FrmStatus.setMaxValue silkDC.Count
  Do While Not silkDC.IsEmpty()
    FrmStatus.setRemaindValue silkDC.Count
    RelaxForGUI last_time, 0
    
    r1 = silkDC.Pop() / 2# ' Hole diameter
    x1 = silkDC.Pop() ' Hole center
    y1 = silkDC.Pop() ' Hole center
    x2 = silkDC.Pop() ' Should be zero
    angle = silkDC.Pop()  ' Rotate angle about 0,0
    MOD_x = silkDC.Pop()  ' Rotate offset
    MOD_y = silkDC.Pop()  ' Rotate offset
    x2 = Abs(x2 - x1)
    
    If (x2 >= 0.005 * InToMeter) Then
      Rotate2D x1, y1, angle, MOD_x, MOD_y
      mySketch.CreateCircleByRadius x1, y1, z1, x2 + r1
      x2 = x2 - r1
      If (x2 >= 0.005 * InToMeter) And (r1 >= 0.005 * InToMeter) Then
        mySketch.CreateCircleByRadius x1, y1, z1, x2
      End If
    End If
  Loop

End Sub

Sub SketchText(Part As IPartDoc, text As Stack_String, text_info As Stack_Dbl, z1 As Double)
  Dim tc, ts, flip
  Dim last_time As Long
    
  Dim x1 As Double, y1 As Double
  Dim x2 As Double, y2 As Double
  Dim MOD_x As Double, MOD_y As Double
  Dim MOD_angle As Double, angle As Double
  Dim s As String, s1 As String

  FrmStatus.setMaxValue text.Count
  
  Do While Not text.IsEmpty()
    FrmStatus.setRemaindValue text.Count
    RelaxForGUI last_time, 0
    
    x1 = text_info.Pop()
    y1 = text_info.Pop()
    x2 = text_info.Pop() / 1.1
    y2 = text_info.Pop() / 1.1
    angle = text_info.Pop()
    MOD_x = text_info.Pop()
    MOD_y = text_info.Pop()
    MOD_angle = text_info.Pop()
    s = text.Pop()
    
    If x2 > 0.0007 And y2 > 0.0007 And s <> "" Then
      Dim skText As ISketchText
      Dim Fmt As ITextFormat
      Dim swLine As ISketchLine
      Dim i
      
      If z1 > 0 Then ' Bottom Layer
        flip = 0
        angle = 180 - angle
        tc = Cos(angle / 180 * PI) * x2
        ts = Sin(angle / 180 * PI) * x2
        Rotate2D x1, y1, MOD_angle, MOD_x - tc * Len(s) / 2, MOD_y - ts * Len(s) / 2
        
        MOD_y = -y2 / 2
      Else
        flip = 1
        angle = 180 + angle
        tc = -Cos(angle / 180 * PI) * x2
        ts = Sin(angle / 180 * PI) * x2
        Rotate2D x1, y1, MOD_angle, MOD_x - tc * Len(s) / 2, MOD_y - ts * Len(s) / 2
        
        MOD_y = y2 / 2
      End If
      MOD_x = x1 - MOD_y * ts / x2
      MOD_y = y1 + MOD_y * tc / x2
      
      s1 = "<r" + Trim(Str(angle)) + ">"
      x2 = x2 / 1.2
      For i = 1 To Len(s)
        Set skText = Part.InsertSketchText(MOD_x, MOD_y, z1, s1 + Mid(s, i, 1) + "</r>", 1, 0, flip, 90, 0)
        Set Fmt = skText.IGetTextFormat()
        Fmt.CharHeight = y2
        Fmt.LineLength = x2
        skText.ISetTextFormat False, Fmt
        MOD_x = MOD_x + tc
        MOD_y = MOD_y + ts
      Next i
    End If
  Loop

End Sub

Sub GenerateVMRL(Part As IPartDoc, FileName As String, _
  Scale_x, Scale_y, Scale_z, _
  Optional genMinMaxBox As Boolean = True, _
  Optional genWireFrame As Boolean = False)
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
  
  Dim mySketch As SketchManager
  Dim myfeature As FeatureManager
  Set mySketch = Part.SketchManager
  mySketch.AddToDB = True
  mySketch.DisplayWhenAdded = False
  
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
      RelaxForGUI last_time, 0
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
              mySketch.Insert3DSketch True
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
              mySketch.InsertSketch True
              mySketch.Create3PointCornerRectangle min_x, min_y, 0#, min_x, max_y, 0#, max_x, max_y, 0#
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
  mySketch.DisplayWhenAdded = True
  mySketch.AddToDB = False
  Part.Extension.AddComment FileName
End Sub

Sub GeneratePCBAssembly(Part As IAssemblyDoc, boardFilename As String, _
  PosFileName As String, BOMFileName As String, _
  Optional overwriteGeneratedVRML = False, _
  Optional genMinMaxBox As Boolean = True, _
  Optional genWireFrame As Boolean = False, _
  Optional RenameComponents As Boolean = True)
  
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
    RelaxForGUI last_time, 0
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
      RelaxForGUI last_time, 0
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
    
    DrillScale = InchToSW / 10000#  ' unit/in
    GerbScale = InchToSW / 10000#   ' unit/in
    
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
    
    PCBCfgForm.txtDrillScale = CStr(InchToSW / DrillScale)
    PCBCfgForm.txtGerbScale = CStr(InchToSW / GerbScale)
    PCBCfgForm.txtPosScale = CStr(InchToSW / POSScale)
    PCBCfgForm.txtPosAngleScale = CStr(AngScale)
    PCBCfgForm.txtWRLScale = CStr(InchToSW / VRMLScale)
    PCBCfgForm.txtPCBThickness = CStr(PCB_Thickness * 1000)
  End If
  PCBCfgForm.Show
End Sub
