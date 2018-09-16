Attribute VB_Name = "KiCad_to_SolidWork"
'
'    PCBFab2Solidwork
'    Copyright (C) 2018  NhatKhai L. Nguyen
'
'    Please check LICENSE file for detail.
'
Option Explicit

Public swApp As ISldWorks

Const IncToMeter = 25.4 / 1000#                 ' (mm/in)

Public KiCadScale As Double     ' =25.4/10000000# Convert Kicad to mm
Public Shape3DScale As Double   ' =25.4 / 1000# Convert Kicad 3D scaling

Public POSScale As Double       '= 1000000 / 25.4 ' mil/unit

Public VRMLScale As Double      ' = 25.4 / 10000# Convert WRL file to mm
Public BOMScale As Double       '= 1000 / 25.4    ' mil/unit

Public PCB_Thickness As Double  ' = 0.063 (inches)

Const CopperTopLayer = 15
Const CopperBotLayer = 0

Const silkTopLayer = 21
Const silkBottomLayer = 20
Const PCBEdgeLayer = 28

Const SilkColorR = 0#
Const SilkColorG = 0#
Const SilkColorB = 0#

Const PCBColorR = 0#
Const PCBColorG = 1#
Const PCBColorB = 0#

Const CopperColorR = 1#
Const CopperColorG = 0#
Const CopperColorB = 0#


Sub GeneratePCB(Part As IPartDoc, FileName As String, _
  Optional gen_silks As Boolean = False, Optional silkDS_SketchOnly As Boolean = False, _
  Optional gen_MinMaxSilks As Boolean = False, _
  Optional gen_textsilk As Boolean = False, _
  Optional filter_minSize As Double = 0#)
  
  Dim inFile As Integer
  Dim st As Stack_String
  Dim silkMinMax_Top As Stack_Dbl
  Dim silkMinMax_Bot As Stack_Dbl
  Dim silk_Top As Stack_Dbl
  Dim silk_Bot As Stack_Dbl
  Dim silkDC_Top As Stack_Dbl
  Dim silkDC_Bot As Stack_Dbl
  
  Dim textCopperTop As Stack_String
  Dim textCopperTop_info As Stack_Dbl
  Dim textCopperBot As Stack_String
  Dim textCopperBot_info As Stack_Dbl
  Dim textTop As Stack_String
  Dim textTop_info As Stack_Dbl
  Dim textBot As Stack_String
  Dim textBot_info As Stack_Dbl
  
  Dim data1() As String, data2() As String, data3() As String, data4() As String
  
  Dim x1 As Double, y1 As Double, z1 As Double
  Dim x2 As Double, y2 As Double, width As Double
  Dim tx1 As Double, ty1 As Double
  Dim tx2 As Double, ty2 As Double
  Dim min_x1 As Double, min_y1 As Double
  Dim max_x1 As Double, max_y1 As Double
  Dim min_x2 As Double, min_y2 As Double
  Dim max_x2 As Double, max_y2 As Double
  Dim min_r2 As Double
  Dim min_r1 As Double
  Dim r1 As Double
  
  Dim shape As Integer, layer As Integer
  Dim angle As Double, mod_x As Double, mod_y As Double, mod_angle As Double
  Dim mat
  
  Dim line As Long, module_name As String
  Dim s As String
  Dim a() As String
  Dim last_time As Long
  Dim mySketch As SketchManager
  Dim myfeature As FeatureManager
  
  inFile = FreeFile
  Open FileName For Input As #inFile
  
  Set st = New Stack_String
  Set silkMinMax_Top = New Stack_Dbl
  Set silkMinMax_Bot = New Stack_Dbl
  Set silk_Top = New Stack_Dbl
  Set silk_Bot = New Stack_Dbl
  Set silkDC_Top = New Stack_Dbl
  Set silkDC_Bot = New Stack_Dbl
  Set textTop = New Stack_String
  Set textTop_info = New Stack_Dbl
  Set textBot = New Stack_String
  Set textBot_info = New Stack_Dbl
  Set textCopperTop = New Stack_String
  Set textCopperTop_info = New Stack_Dbl
  Set textCopperBot = New Stack_String
  Set textCopperBot_info = New Stack_Dbl
  
  Part.ClearSelection2 True
  Part.SelectionManager.EnableContourSelection = True
  Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
  
  Set myfeature = Part.FeatureManager
  Set mySketch = Part.SketchManager
  mySketch.AddToDB = True
  mySketch.DisplayWhenAdded = False
  mySketch.InsertSketch True
  mySketch.ActiveSketch.Name = "Sketch_PCBBoard"
  
  filter_minSize = filter_minSize * IncToMeter
  
  FrmStatus.AppendTODO "Read PCB File " + FileName
  line = 0
  Do While Not EOF(inFile)
    RelaxForGUI last_time, 0
    Line Input #inFile, s
    line = line + 1
    a = Split(s)
    If LBound(a) <= UBound(a) Then
      Select Case a(0)
        Case "$DRAWSEGMENT"
          shape = -1
          st.Push a(0)
        Case "$EndDRAWSEGMENT"
          Do While st.Pop() <> "$DRAWSEGMENT"
          Loop
          
        Case "$MODULE"
          shape = 0
          min_x1 = 1000
          max_x1 = -1000
          min_y1 = 1000
          max_y1 = -1000
          min_x2 = 1000
          max_x2 = -1000
          min_y2 = 1000
          max_y2 = -1000
          min_r1 = 1000
          min_r2 = 1000
          module_name = a(1)
          st.Push a(0)
          
          textTop.Save
          textTop_info.Save
          silkDC_Top.Save
          silk_Top.Save
          textBot.Save
          textBot_info.Save
          silkDC_Bot.Save
          silk_Bot.Save
        Case "$EndMODULE"
          x2 = Abs(max_x1 - min_x1) * (max_y1 - min_y1)
          y2 = Abs(max_x2 - min_x2) * (max_y2 - min_y2)
          If (x2 < filter_minSize) And (y2 < filter_minSize) Then
            textTop.Restore
            textTop_info.Restore
            silkDC_Top.Restore
            silk_Top.Restore
            textBot.Restore
            textBot_info.Restore
            silkDC_Bot.Restore
            silk_Bot.Restore
          Else
            If gen_MinMaxSilks Then
              If (max_x1 - min_x1 > 0.01 * IncToMeter) _
              And (max_y1 - min_y1 > 0.01 * IncToMeter) Then
                silkMinMax_Top.Push x1
                silkMinMax_Top.Push y1
                silkMinMax_Top.Push angle
                silkMinMax_Top.Push min_r1
                silkMinMax_Top.Push max_x1
                silkMinMax_Top.Push max_y1
                silkMinMax_Top.Push min_x1
                silkMinMax_Top.Push min_y1
              End If
              
              If (max_x2 - min_x2 > 0.01 * IncToMeter) _
              And (max_y2 - min_y2 > 0.01 * IncToMeter) Then
                silkMinMax_Bot.Push x1
                silkMinMax_Bot.Push y1
                silkMinMax_Bot.Push angle
                silkMinMax_Bot.Push min_r2
                silkMinMax_Bot.Push max_x2
                silkMinMax_Bot.Push max_y2
                silkMinMax_Bot.Push min_x2
                silkMinMax_Bot.Push min_y2
              End If
            End If
          End If
          
          shape = 0
          Do While st.Pop() <> "$MODULE"
          Loop
          
        Case "$PAD"
          shape = shape And 1
          st.Push a(0)
        Case "$EndPAD"
          If (shape = 7) And (r1 >= 0.005 * IncToMeter) Then
            Rotate2D x2, y2, angle, x1, y1
            r1 = r1 / 2
            mySketch.CreateCircleByRadius x2, y2, 0#, r1
          End If
          shape = shape And 1
          Do While st.Pop() <> "$PAD"
          Loop
          
        Case "$TEXTPCB"
          shape = -1
          st.Push a(0)
          ReDim data1(1)
          ReDim data2(0)
          ReDim data3(0)
        Case "$EndTEXTPCB"
          If (data1(0) <> "") And (UBound(data2) > 0) And (UBound(data3) > 0) Then
            Select Case CInt(data3(1))
              Case silkTopLayer
                textTop.Push data1(0)
                textTop_info.Push 0#  ' Module Angle
                textTop_info.Push CDbl(data2(2)) * KiCadScale ' y1
                textTop_info.Push CDbl(data2(1)) * KiCadScale ' x1
                textTop_info.Push CDbl(data2(6)) / 10 ' Text Angle
                textTop_info.Push CDbl(data2(4)) * KiCadScale ' Text ySize
                textTop_info.Push CDbl(data2(3)) * KiCadScale ' Text xSize
                textTop_info.Push 0#
                textTop_info.Push 0#
                
              Case silkBottomLayer
                textBot.Push data1(0)
                textBot_info.Push 0#  ' Module Angle
                textBot_info.Push CDbl(data2(2)) * KiCadScale ' y1
                textBot_info.Push CDbl(data2(1)) * KiCadScale ' x1
                textBot_info.Push CDbl(data2(6)) / 10 ' Text Angle
                textBot_info.Push CDbl(data2(4)) * KiCadScale ' Text ySize
                textBot_info.Push CDbl(data2(3)) * KiCadScale ' Text xSize
                textBot_info.Push 0#
                textBot_info.Push 0#
                
              Case CopperTopLayer
                textCopperTop.Push data1(0)
                textCopperTop_info.Push 0#  ' Module Angle
                textCopperTop_info.Push CDbl(data2(2)) * KiCadScale ' y1
                textCopperTop_info.Push CDbl(data2(1)) * KiCadScale ' x1
                textCopperTop_info.Push CDbl(data2(6)) / 10 ' Text Angle
                textCopperTop_info.Push CDbl(data2(4)) * KiCadScale ' Text ySize
                textCopperTop_info.Push CDbl(data2(3)) * KiCadScale ' Text xSize
                textCopperTop_info.Push 0#
                textCopperTop_info.Push 0#
                
              Case CopperBotLayer
                textCopperBot.Push data1(0)
                textCopperBot_info.Push 0#  ' Module Angle
                textCopperBot_info.Push CDbl(data2(2)) * KiCadScale ' y1
                textCopperBot_info.Push CDbl(data2(1)) * KiCadScale ' x1
                textCopperBot_info.Push CDbl(data2(6)) / 10 ' Text Angle
                textCopperBot_info.Push CDbl(data2(4)) * KiCadScale ' Text ySize
                textCopperBot_info.Push CDbl(data2(3)) * KiCadScale ' Text xSize
                textCopperBot_info.Push 0#
                textCopperBot_info.Push 0#
            End Select
          End If
          Do While st.Pop() <> "$TEXTPCB"
          Loop
          
        Case Else
          Select Case st.Top()
            Case "$DRAWSEGMENT"
              If a(0) = "Po" Then
                shape = CInt(a(1))
                x1 = CDbl(a(2)) * KiCadScale
                y1 = CDbl(a(3)) * KiCadScale
                x2 = CDbl(a(4)) * KiCadScale
                y2 = CDbl(a(5)) * KiCadScale
                width = CDbl(a(6)) * KiCadScale
              ElseIf a(0) = "De" Then
                ' PCB Edge Layer
                If CInt(a(1)) = PCBEdgeLayer Then
                  Select Case shape
                    Case 0  ' Line
                      mySketch.CreateLine x1, y1, 0#, x2, y2, 0#
                    Case 2  ' Arc
                      mySketch.CreateArc x1, y1, 0#, x2, y2, 0#, x1 + y1 - y2, y1 + x2 - x1, 0#, 1
                    Case 3  ' Circle
                      mySketch.CreateCircle x1, y1, 0#, x2, y2, 0#
                    Case Else
                      MsgBox "Shape Not Supported" + s + "@Line " + Str(line)
                  End Select
                Else
                  Select Case shape
                    Case 0  ' Line
                      Select Case CInt(a(1))
                        Case silkTopLayer
                          silk_Top.Push 0
                          silk_Top.Push 0
                          silk_Top.Push 0
                          silk_Top.Push y2
                          silk_Top.Push x2
                          silk_Top.Push y1
                          silk_Top.Push x1
                          silk_Top.Push width
                          
                        Case silkBottomLayer
                          silk_Bot.Push 0
                          silk_Bot.Push 0
                          silk_Bot.Push 0
                          silk_Bot.Push y2
                          silk_Bot.Push x2
                          silk_Bot.Push y1
                          silk_Bot.Push x1
                      End Select
                      
                    Case 2  ' Arc
                      ' TODO
                      
                    Case 3  ' Circle
                      Select Case CInt(a(1))
                        Case silkTopLayer
                          silkDC_Top.Push 0
                          silkDC_Top.Push 0
                          silkDC_Top.Push 0
                          silkDC_Top.Push x1 + Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2) ' Radius
                          silkDC_Top.Push y1
                          silkDC_Top.Push x1
                          silkDC_Top.Push width
                        
                        Case silkBottomLayer
                          silkDC_Bot.Push 0
                          silkDC_Bot.Push 0
                          silkDC_Bot.Push 0
                          silkDC_Bot.Push x1 + Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2) ' Radius
                          silkDC_Bot.Push y1
                          silkDC_Bot.Push x1
                          silkDC_Bot.Push width
                      End Select
                      
                    Case Else
                      MsgBox "Shape Not Supported" + s + "@Line " + Str(line)
                  End Select
                End If
              End If
              
            Case "$MODULE"
              Select Case a(0)
                Case "Po"
                  If ((shape And 1) = 0) Then
                    x1 = CDbl(a(1)) * KiCadScale
                    y1 = CDbl(a(2)) * KiCadScale
                    angle = CDbl(a(3)) / 10
                    shape = shape Or 1
                  End If
                  
                Case "DS"
                  If shape And 1 <> 0 Then
                    ty2 = CDbl(a(4)) * KiCadScale ' y2
                    tx2 = CDbl(a(3)) * KiCadScale ' x2
                    ty1 = CDbl(a(2)) * KiCadScale ' y1
                    tx1 = CDbl(a(1)) * KiCadScale ' x1
                    r1 = CDbl(a(5)) * KiCadScale
                  
                    Select Case CInt(a(6))
                      Case silkTopLayer
                        silk_Top.Push y1
                        silk_Top.Push x1
                        silk_Top.Push angle
                        silk_Top.Push ty2 ' y2
                        silk_Top.Push tx2 ' x2
                        silk_Top.Push ty1 ' y1
                        silk_Top.Push tx1 ' x1
                        silk_Top.Push r1 ' Width
                        If min_x1 > tx1 Then min_x1 = tx1
                        If max_x1 < tx1 Then max_x1 = tx1
                        If min_x1 > tx2 Then min_x1 = tx2
                        If max_x1 < tx2 Then max_x1 = tx2
                        
                        If min_y1 > ty1 Then min_y1 = ty1
                        If max_y1 < ty1 Then max_y1 = ty1
                        If min_y1 > ty2 Then min_y1 = ty2
                        If max_y1 < ty2 Then max_y1 = ty2
                        
                        If min_r1 > r1 Then min_r1 = r1
                        
                      Case silkBottomLayer
                        silk_Bot.Push y1
                        silk_Bot.Push x1
                        silk_Bot.Push angle
                        silk_Bot.Push ty2 ' y2
                        silk_Bot.Push tx2 ' x2
                        silk_Bot.Push ty1 ' y1
                        silk_Bot.Push tx1 ' x1
                        silk_Bot.Push r1 ' Width
                        If min_x2 > tx1 Then min_x2 = tx1
                        If max_x2 < tx1 Then max_x2 = tx1
                        If min_x2 > tx2 Then min_x2 = tx2
                        If max_x2 < tx2 Then max_x2 = tx2
      
                        If min_y2 > ty1 Then min_y2 = ty1
                        If max_y2 < ty1 Then max_y2 = ty1
                        If min_y2 > ty2 Then min_y2 = ty2
                        If max_y2 < ty2 Then max_y2 = ty2
                        
                        If min_r2 > r1 Then min_r2 = r1

                    End Select
                  End If
                  
                Case "DC"
                  If shape And 1 <> 0 Then
                    Select Case CInt(a(6))
                      Case silkTopLayer
                        silkDC_Top.Push y1
                        silkDC_Top.Push x1
                        silkDC_Top.Push angle
                        silkDC_Top.Push CDbl(a(3)) * KiCadScale ' radius
                        silkDC_Top.Push CDbl(a(2)) * KiCadScale ' y1
                        silkDC_Top.Push CDbl(a(1)) * KiCadScale ' x1
                        silkDC_Top.Push CDbl(a(5)) * KiCadScale ' Width
                      
                      Case silkBottomLayer
                        silkDC_Bot.Push y1
                        silkDC_Bot.Push x1
                        silkDC_Bot.Push angle
                        silkDC_Bot.Push CDbl(a(3)) * KiCadScale ' radius
                        silkDC_Bot.Push CDbl(a(2)) * KiCadScale ' y1
                        silkDC_Bot.Push CDbl(a(1)) * KiCadScale ' x1
                        silkDC_Bot.Push CDbl(a(5)) * KiCadScale ' Width
                    End Select
                  End If
                  
                Case "T0", "T1"
                  If a(8) = "V" And (shape And 1 <> 0) Then
                    s = StrJoin(a, 11, UBound(a))
                    x2 = CDbl(a(5)) / 10 ' Angle
                    Do While x2 > 90
                      x2 = x2 - 180
                    Loop
                    Do While x2 < 0
                      x2 = x2 + 180
                    Loop
      
                    
                    Select Case CInt(a(9))
                      Case silkTopLayer
                        textTop.Push Mid(s, 2, Len(s) - 2)
                        textTop_info.Push angle
                        textTop_info.Push y1
                        textTop_info.Push x1
                        textTop_info.Push x2 ' Angle
                        textTop_info.Push CDbl(a(4)) * KiCadScale ' Text ySize
                        textTop_info.Push CDbl(a(3)) * KiCadScale ' Text xSize
                        textTop_info.Push CDbl(a(2)) * KiCadScale ' y1
                        textTop_info.Push CDbl(a(1)) * KiCadScale ' x1
                        
                      Case silkBottomLayer
                        textBot.Push Mid(s, 2, Len(s) - 2)
                        textBot_info.Push angle
                        textBot_info.Push y1
                        textBot_info.Push x1
                        textBot_info.Push x2 ' Angle
                        textBot_info.Push CDbl(a(4)) * KiCadScale ' Text ySize
                        textBot_info.Push CDbl(a(3)) * KiCadScale ' Text xSize
                        textBot_info.Push CDbl(a(2)) * KiCadScale ' y1
                        textBot_info.Push CDbl(a(1)) * KiCadScale ' x1
                    End Select
                    
                  End If
              End Select
              
            Case "$PAD"
              If (a(0) = "Po") And ((shape And 2) = 0) Then
                x2 = CDbl(a(1)) * KiCadScale
                y2 = CDbl(a(2)) * KiCadScale
                shape = shape Or 2
              ElseIf (a(0) = "Dr") And ((shape And 4) = 0) Then
                r1 = CDbl(a(1)) * KiCadScale
                shape = shape Or 4
              End If
              
            Case "$TEXTPCB"
              Select Case a(0)
                Case "Te"
                  s = StrJoin(a, 1, UBound(a))
                  data1(0) = Mid(s, 2, Len(s) - 2)
                Case "nl"
                  s = StrJoin(a, 1, UBound(a))
                  data1(0) = data1(0) + vbCr + Mid(s, 2, Len(s) - 2)
                Case "Po"
                  data2 = a
                Case "De"
                  data3 = a
              End Select
              
          End Select ' st.Top()
          
      End Select ' s
    End If
  Loop
  
  Close #inFile
  
  Dim feature As IFeature, tf
  
  Set feature = myfeature.FeatureExtrusion2( _
    True, False, False, _
    swEndCondMidPlane, swEndCondMidPlane, _
    PCB_Thickness * IncToMeter, PCB_Thickness * IncToMeter, _
    False, False, False, False, _
    0.01745329251994, 0.01745329251994, _
    False, False, False, False, _
    True, True, True, _
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
  
  FrmStatus.AppendTODO "Generate " + Str(silk_Bot.Count / 8) + " Component Silks on Bottom Layer"
  FrmStatus.AppendTODO "Generate " + Str(silk_Top.Count / 8) + " Component Silks on Top Layer"
  FrmStatus.AppendTODO "Generate " + Str(silkMinMax_Bot.Count / 8) + " Component MinMax Silks on Bottom Layer"
  FrmStatus.AppendTODO "Generate " + Str(silkMinMax_Top.Count / 8) + " Component MinMax Silks on Top Layer"
  FrmStatus.AppendTODO "Generate " + Str(silkDC_Bot.Count / 7) + " Circle Silks on Bottom Layer"
  FrmStatus.AppendTODO "Generate " + Str(silkDC_Top.Count / 7) + " Circle Silks on Top Layer"
  FrmStatus.AppendTODO "Generate " + Str(textCopperBot.Count) + " Copper Text on Bottom Layer"
  FrmStatus.AppendTODO "Generate " + Str(textCopperTop.Count) + " Copper Text on Top Layer"
  FrmStatus.AppendTODO "Generate " + Str(textBot.Count) + " Text Silks on Bottom Layer"
  FrmStatus.AppendTODO "Generate " + Str(textTop.Count) + " Text Silks on Top Layer"
  
  'Generate Silk Texts
  If gen_textsilk Then
    z1 = -(PCB_Thickness / 2 + 0.001) * IncToMeter
    
    If textTop.Count > 0 Then
      Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
      mySketch.InsertSketch True
      mySketch.ActiveSketch.Name = "Sketch_SilkText_Top"
      SketchText Part, textTop, textTop_info, z1
      Set feature = myfeature.FeatureExtrusion2( _
        True, False, False, _
        swEndCondMidPlane, swEndCondMidPlane, _
        0.001 * IncToMeter, 0.001 * IncToMeter, _
        False, False, False, False, _
        0.01745329251994, 0.01745329251994, _
        False, False, False, False, _
        True, True, True, _
        swStartOffset, z1, False)
      If Not feature Is Nothing Then
        feature.Name = "SilkTexts_Top"
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
    FrmStatus.PopTODO
    
    If textBot.Count > 0 Then
      Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
      mySketch.InsertSketch True
      mySketch.ActiveSketch.Name = "Sketch_SilkText_Bot"
      SketchText Part, textBot, textBot_info, -z1
      Set feature = myfeature.FeatureExtrusion2( _
        True, False, False, _
        swEndCondMidPlane, swEndCondMidPlane, _
        0.001 * IncToMeter, 0.001 * IncToMeter, _
        False, False, False, False, _
        0.01745329251994, 0.01745329251994, _
        False, False, False, False, _
        True, True, True, _
        swStartOffset, -z1, False)
      If Not feature Is Nothing Then
        feature.Name = "SilkTexts_Bot"
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
    FrmStatus.PopTODO
  Else
    FrmStatus.PopTODO
    FrmStatus.PopTODO
  End If
  
  ' Generate Copper Texts
  If gen_textsilk Then
    z1 = -(PCB_Thickness / 2 + 0.001) * IncToMeter
    
    If textCopperTop.Count > 0 Then
      Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
      mySketch.InsertSketch True
      mySketch.ActiveSketch.Name = "Sketch_CopperText_Top"
      SketchText Part, textCopperTop, textCopperTop_info, z1
      Set feature = myfeature.FeatureExtrusion2( _
        True, False, False, _
        swEndCondMidPlane, swEndCondMidPlane, _
        0.001 * IncToMeter, 0.001 * IncToMeter, _
        False, False, False, False, _
        0.01745329251994, 0.01745329251994, _
        False, False, False, False, _
        True, True, True, _
        swStartOffset, z1, False)
      If Not feature Is Nothing Then
        feature.Name = "CopperTexts_Top"
        mat = feature.GetMaterialPropertyValues2(swThisConfiguration, Nothing)
        mat(0) = CopperColorR
        mat(1) = CopperColorG
        mat(2) = CopperColorB
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
    FrmStatus.PopTODO
    
    If textCopperBot.Count > 0 Then
      Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
      mySketch.InsertSketch True
      mySketch.ActiveSketch.Name = "Sketch_CopperText_Bot"
      SketchText Part, textCopperBot, textCopperBot_info, -z1
      Set feature = myfeature.FeatureExtrusion2( _
        True, False, False, _
        swEndCondMidPlane, swEndCondMidPlane, _
        0.001 * IncToMeter, 0.001 * IncToMeter, _
        False, False, False, False, _
        0.01745329251994, 0.01745329251994, _
        False, False, False, False, _
        True, True, True, _
        swStartOffset, -z1, False)
      If Not feature Is Nothing Then
        feature.Name = "CopperTexts_Bot"
        mat = feature.GetMaterialPropertyValues2(swThisConfiguration, Nothing)
        mat(0) = CopperColorR
        mat(1) = CopperColorG
        mat(2) = CopperColorB
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
    FrmStatus.PopTODO
  Else
    FrmStatus.PopTODO
    FrmStatus.PopTODO
  End If
    
  ' Generate Circle Silks normall good for polarity indicators
  If gen_silks Or gen_textsilk Then
    z1 = -(PCB_Thickness / 2 + 0.001) * IncToMeter
    
    If silkDC_Top.Count > 0 Then
      Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
      mySketch.InsertSketch True
      mySketch.ActiveSketch.Name = "Sketch_SilkDC_Top"
      SketchDCSilk mySketch, silkDC_Top, z1
      Set feature = myfeature.FeatureExtrusion2( _
        True, False, False, _
        swEndCondMidPlane, swEndCondMidPlane, _
        0.001 * IncToMeter, 0.001 * IncToMeter, _
        False, False, False, False, _
        0.01745329251994, 0.01745329251994, _
        False, False, False, False, _
        True, True, True, _
        swStartOffset, z1, False)
      If Not feature Is Nothing Then
        feature.Name = "SilkDC_Top"
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
    FrmStatus.PopTODO
    
    If silkDC_Bot.Count > 0 Then
      Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
      mySketch.InsertSketch True
      mySketch.ActiveSketch.Name = "Sketch_SilkDC_Bot"
      SketchDCSilk mySketch, silkDC_Bot, -z1
      Set feature = myfeature.FeatureExtrusion2( _
        True, False, False, _
        swEndCondMidPlane, swEndCondMidPlane, _
        0.001 * IncToMeter, 0.001 * IncToMeter, _
        False, False, False, False, _
        0.01745329251994, 0.01745329251994, _
        False, False, False, False, _
        True, True, True, _
        swStartOffset, -z1, False)
      If Not feature Is Nothing Then
        feature.Name = "SilkDC_Bot"
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
    FrmStatus.PopTODO
  Else
    FrmStatus.PopTODO
    FrmStatus.PopTODO
  End If
    
  ' Generate Boundary Skils for Components
  If gen_MinMaxSilks Then
    z1 = -(PCB_Thickness / 2 + 0.001) * IncToMeter
    
    If silkMinMax_Top.Count > 0 Then
      Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
      mySketch.InsertSketch True
      mySketch.ActiveSketch.Name = "Sketch_SilkMinMax_Top"
      SketchRecSilk mySketch, silkMinMax_Top, z1
      Set feature = myfeature.FeatureExtrusion2( _
        True, False, False, _
        swEndCondMidPlane, swEndCondMidPlane, _
        0.001 * IncToMeter, 0.001 * IncToMeter, _
        False, False, False, False, _
        0.01745329251994, 0.01745329251994, _
        False, False, False, False, _
        True, True, True, _
        swStartOffset, z1, False)
      If Not feature Is Nothing Then
        feature.Name = "SilkMinMax_Top"
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
    FrmStatus.PopTODO
    
    If silkMinMax_Bot.Count > 0 Then
      Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
      mySketch.InsertSketch True
      mySketch.ActiveSketch.Name = "Sketch_SilkMinMax_Bot"
      SketchRecSilk mySketch, silkMinMax_Bot, -z1
      Set feature = myfeature.FeatureExtrusion2( _
        True, False, False, _
        swEndCondMidPlane, swEndCondMidPlane, _
        0.001 * IncToMeter, 0.001 * IncToMeter, _
        False, False, False, False, _
        0.01745329251994, 0.01745329251994, _
        False, False, False, False, _
        True, True, True, _
        swStartOffset, -z1, False)
      If Not feature Is Nothing Then
        feature.Name = "SilkMinMax_Bot"
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
    FrmStatus.PopTODO
  Else
    FrmStatus.PopTODO
    FrmStatus.PopTODO
  End If
  
  ' Generate Detail Silk for Components (Took alot of time and memory)
  If gen_silks Then
    z1 = -(PCB_Thickness / 2 + 0.001) * IncToMeter
    
    If silk_Top.Count > 0 Then
      Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
      mySketch.InsertSketch True
      mySketch.ActiveSketch.Name = "Sketch_SilkDS_Top"
      If Not silkDS_SketchOnly Then
        SketchDSSilk mySketch, silk_Top, z1, myfeature, "SilkDS_Top"
      Else
        SketchDSSilk mySketch, silk_Top, z1
        mySketch.InsertSketch True
      End If
    End If
    FrmStatus.PopTODO
    
    If silk_Bot.Count > 0 Then
      Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
      mySketch.InsertSketch True
      mySketch.ActiveSketch.Name = "Sketch_SilkDS_Bot"
      If Not silkDS_SketchOnly Then
        SketchDSSilk mySketch, silk_Bot, -z1, myfeature, "SilkDS_Bot"
      Else
        SketchDSSilk mySketch, silk_Bot, -z1
        mySketch.InsertSketch True
      End If
      mySketch.InsertSketch True
    End If
    FrmStatus.PopTODO
  Else
    FrmStatus.PopTODO
    FrmStatus.PopTODO
  End If
  
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
  Dim mod_x As Double, mod_y As Double
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
    mod_y = silks.Pop()
    mod_x = silks.Pop()
    
    If (r1 < 0.008 * IncToMeter) Then
      r1 = 0.008 * IncToMeter
    End If
    x3 = x1 + r1
    y3 = y1 + r1
    x4 = x2 - r1
    y4 = y2 - r1
    If ((x4 - x3 < 0.005 * IncToMeter) Or (y4 - y3 < 0.005 * IncToMeter)) Then
      r1 = 0
    End If
      
    Rotate2D x1, y1, angle, mod_x, mod_y
    Rotate2D x2, y2, angle, mod_x, mod_y
    Rotate2D x3, y3, angle, mod_x, mod_y
    Rotate2D x4, y4, angle, mod_x, mod_y
    
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
  Dim mod_x As Double, mod_y As Double
  Dim mod_angle As Double, angle As Double
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
    mod_x = silks.Pop()
    mod_y = silks.Pop()
    
    If (Abs(x1 - x2) > r1 Or Abs(y1 - y2) > r1) And (r1 >= 0.005 * IncToMeter) Then
      Rotate2D x1, y1, angle, mod_x, mod_y
      Rotate2D x2, y2, angle, mod_x, mod_y
      
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
          0.001 * IncToMeter, 0.001 * IncToMeter, _
          False, False, False, False, _
          0.01745329251994, 0.01745329251994, _
          False, False, False, False, _
          True, True, True, _
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
  Dim mod_x As Double, mod_y As Double
  Dim mod_angle As Double, angle As Double
  Dim last_time As Long

  FrmStatus.setMaxValue silkDC.Count
  Do While Not silkDC.IsEmpty()
    FrmStatus.setRemaindValue silkDC.Count
    RelaxForGUI last_time, 0
    
    r1 = silkDC.Pop() / 2#
    x1 = silkDC.Pop()
    y1 = silkDC.Pop()
    x2 = silkDC.Pop()
    angle = silkDC.Pop()
    mod_x = silkDC.Pop()
    mod_y = silkDC.Pop()
    x2 = Abs(x2 - x1)
    
    If (x2 >= 0.005 * IncToMeter) Then
      Rotate2D x1, y1, angle, mod_x, mod_y
      mySketch.CreateCircleByRadius x1, y1, z1, x2 + r1
      x2 = x2 - r1
      If (x2 >= 0.005 * IncToMeter) And (r1 >= 0.005 * IncToMeter) Then
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
  Dim mod_x As Double, mod_y As Double
  Dim mod_angle As Double, angle As Double
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
    mod_x = text_info.Pop()
    mod_y = text_info.Pop()
    mod_angle = text_info.Pop()
    s = text.Pop()
    
    If x2 > 0.0007 And y2 > 0.0007 And s <> "" Then
      Dim skText As ISketchText
      Dim fmt As ITextFormat
      Dim swLine As ISketchLine
      Dim i
      
      If z1 > 0 Then ' Bottom Layer
        flip = 0
        angle = 180 - angle
        tc = Cos(angle / 180 * PI) * x2
        ts = Sin(angle / 180 * PI) * x2
        Rotate2D x1, y1, mod_angle, mod_x - tc * Len(s) / 2, mod_y - ts * Len(s) / 2
        
        mod_y = -y2 / 2
      Else
        flip = 1
        angle = 180 + angle
        tc = -Cos(angle / 180 * PI) * x2
        ts = Sin(angle / 180 * PI) * x2
        Rotate2D x1, y1, mod_angle, mod_x - tc * Len(s) / 2, mod_y - ts * Len(s) / 2
        
        mod_y = y2 / 2
      End If
      mod_x = x1 - mod_y * ts / x2
      mod_y = y1 + mod_y * tc / x2
      
      s1 = "<r" + Trim(Str(angle)) + ">"
      x2 = x2 / 1.2
      For i = 1 To Len(s)
        Set skText = Part.InsertSketchText(mod_x, mod_y, z1, s1 + Mid(s, i, 1) + "</r>", 1, 0, flip, 90, 0)
        Set fmt = skText.IGetTextFormat()
        fmt.CharHeight = y2
        fmt.LineLength = x2
        skText.ISetTextFormat False, fmt
        mod_x = mod_x + tc
        mod_y = mod_y + ts
      Next i
    End If
  Loop

End Sub

Sub GenerateVMRL(Part As IPartDoc, FileName As String, scale_x, scale_y, scale_z, _
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
            If ((max_z - min_z) > 0.001 * IncToMeter) _
               And ((max_x - min_x) > 0.001 * IncToMeter) _
               And ((max_y - min_y) > 0.001 * IncToMeter) _
            Then
              Part.ClearSelection2 True
              Part.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
              mySketch.InsertSketch True
              mySketch.Create3PointCornerRectangle min_x, min_y, 0, min_x, max_y, 0#, max_x, max_y, 0#
              Set pp = Part.FeatureManager.FeatureExtrusion2( _
                True, False, False, _
                swEndCondMidPlane, swEndCondMidPlane, _
                (max_z - min_z), (max_z - min_z), _
                False, False, False, False, _
                0.01745329251994, 0.01745329251994, _
                False, False, False, False, _
                True, True, True, _
                swStartOffset, (max_z + min_z) / 2, False)
            End If
          End If
          
        Case ""
        
        Case Else
          If st.Top = "point" Then
            If a(i) = "," Then
              If (xi = 3) Then
                x(0) = x(0) * scale_x * VRMLScale
                x(1) = x(1) * scale_y * VRMLScale
                x(2) = x(2) * scale_z * VRMLScale
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
  
  'Dim bodies() As IBody2
  'Dim body As Variant
  'Dim sel As SldWorks.SelectData
  'On Error GoTo DONE_LABEL
  'bodies = Part.GetBodies2(swAllBodies, True)
  'For Each body In bodies
  '  body.Select2 False, sel
  '  Part.FeatureManager.InsertScale swScaleAboutOrigin, False, scale_x * VRMLScale, scale_y * VRMLScale, scale_z * VRMLScale
  'Next body

DONE_LABEL:
  mySketch.DisplayWhenAdded = True
  mySketch.AddToDB = False
  Part.Extension.AddComment FileName
End Sub

Sub GeneratePCBAssembly(Part As IAssemblyDoc, FileName As String, _
  Optional regenerateSLDPRT As Boolean = False, _
  Optional useWRLFirst As Boolean = True, _
  Optional genWireFrame As Boolean = False, _
  Optional filter_minSize As Double = 0, _
  Optional RenameComponents As Boolean = False)
  
  Dim inFile As Integer
  Dim a() As String, s As String
  
  Dim posFile As Integer
  Dim bomFile As Integer
  
  Dim st As Stack_String          ' Stack keep track of read brd state
  Dim compXforms As Stack_Dbl
  Dim compNames As Stack_String
  Dim compRef As Stack_String
  Dim compVal As Stack_String
  
  Dim line, module_name, mod_line, mod_ref, mod_val, mod_layer
  Dim data_flag As Integer
  Dim mod_x, mod_y, mod_angle As Double
  Dim ofs_x, ofs_y, ofs_z, scale_x, scale_y, scale_z, uni_scale As Double
  Dim Rot_x, Rot_y, Rot_z As Double
  Dim Cx, Cy, Cz, Sx, Sy, Sz, Ca, Sa As Double
  Dim Na As String, tempStr As String, myPath As String, tempStr2 As String
  Dim myPart As IPartDoc
  Dim start_time As Date
  
  Dim tx1 As Double, ty1 As Double
  Dim tx2 As Double, ty2 As Double
  Dim min_x As Double, min_y As Double
  Dim max_x As Double, max_y As Double
  
  Dim vcomponents As Variant
  Dim tmpComp As Component2
  Dim skPoint As SketchPoint

  Dim longstatus As Long
  Dim i, j As Long
  Dim boolstatus As Boolean
  Dim last_time As Long
  Dim dict

  Set st = New Stack_String
  Set compNames = New Stack_String
  Set compXforms = New Stack_Dbl
  Set compRef = New Stack_String
  Set compVal = New Stack_String
  Set dict = CreateObject("Scripting.Dictionary")

  
  FrmStatus.AppendTODO "Rename components"
  FrmStatus.AppendTODO "Insert components"
  FrmStatus.AppendTODO "Read PCB File" + FileName
  
  filter_minSize = filter_minSize * IncToMeter
  
  Part.SetAddToDB True
  
  ' Insert PCB
  compNames.Push RemoveFileExt(FileName) + ".sldprt"
  
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
  
  inFile = FreeFile
  Open FileName For Input As #inFile
  posFile = FreeFile
  Open RemoveFileExt(FileName) + ".xyr" For Output As #posFile
  bomFile = FreeFile
  Open RemoveFileExt(FileName) + "_3DBom.csv" For Output As #bomFile
              
  myPath = FilePath(FileName)
  
  Print #posFile, "Ref              Value                      EMPTY  EMPTY  X        Y        Rot     Side"
  Print #bomFile, "#path = " + myPath
  'Print #bomFile, "Ref                      ,Value                    ,Scale_X ,Scale_Y ,Scale_Z ,Ofs_X   ,Ofs_Y   ,Ofs_Z   ,Rot_X   ,Rot_Y   ,Rot_Z   ,3D File"
  Print #bomFile, "Ref,Value,Scale_X,Scale_Y,Scale_Z,Ofs_X,Ofs_Y,Ofs_Z,Rot_X,Rot_Y,Rot_Z,3D File"
  
  line = 0
  start_time = Now
  Do While Not EOF(inFile)
    RelaxForGUI last_time, 0
    Line Input #inFile, s
    line = line + 1
    a = Split(s)
    If LBound(a) <= UBound(a) Then
      Select Case a(0)
        Case "$MODULE"
          data_flag = 0
          module_name = a(1)
          mod_line = line
          mod_ref = ""
          mod_val = ""
          min_x = 1000
          max_x = -1000
          min_y = 1000
          max_y = -1000
          
          compXforms.Save
          compRef.Save
          compVal.Save
          compNames.Save
          
          st.Push a(0)
          
        Case "$EndMODULE"
          data_flag = 0
          If (Abs(max_x - min_x) < filter_minSize) _
            Or (Abs(max_y - min_y) < filter_minSize) Then
            compXforms.Restore
            compRef.Restore
            compVal.Restore
            compNames.Restore
          End If
          Do While st.Pop() <> "$MODULE"
          Loop
          
        Case "$PAD"
          st.Push a(0)
          tx1 = 0
          ty1 = 0
          tx2 = 0
          ty2 = 0
          
        Case "$EndPAD"
          ' Calculate Min Max Rectangular
          tx1 = tx1 - tx2
          ty1 = ty1 - ty2
          tx2 = tx1 + 2 * tx2
          ty2 = ty1 + 2 * ty2
          If min_x > tx1 Then min_x = tx1
          If max_x < tx2 Then max_x = tx2
          If min_y > ty1 Then min_y = ty1
          If max_y < ty2 Then max_y = ty2
          Do While st.Pop() <> "$PAD"
          Loop
          
        Case "$SHAPE3D"
          data_flag = data_flag And 1
          scale_x = 1
          scale_y = 1
          scale_z = 1
          ofs_x = 0
          ofs_y = 0
          ofs_z = 0
          Rot_x = 0
          Rot_y = 0
          Rot_z = 0
          st.Push a(0)
          
        Case "$EndSHAPE3D"
          If (data_flag = 3) Then
            tempStr = RemoveFileExt(Na)
            
            ' Check for uniform scaleing
            If Abs(scale_x - scale_y) < 0.0002 _
            And Abs(scale_x - scale_z) < 0.0002 Then
              tempStr2 = ""
              uni_scale = scale_x
              scale_x = 1
              scale_y = 1
              scale_z = 1
            Else
              tempStr2 = "_" + Str(scale_x) + _
                        "_" + Str(scale_y) + _
                        "_" + Str(scale_z)
              uni_scale = 1
            End If
            
            ' Search for model files
            Dim extList, ext
            If useWRLFirst Then
                extList = Array(".wrl", ".step", ".stp", "")
            Else
                extList = Array(".step", ".stp", ".wrl", "")
            End If
            
            For Each ext In extList
              Na = FindFile(myPath, tempStr, ext)
              If Na <> "" Then
                Exit For
              End If
            Next
                
            ' Check if base solidwork model already exist
            If tempStr2 <> "" Then
              extList = FindFile(myPath, tempStr, ".sldprt")
              If extList <> "" Then
                If FileDateTime(Na + ext) <= FileDateTime(extList + ".sldprt") Then
                  ext = ".sldprt"
                End If
              End If
            End If
            
            ' Generate Position and 3DBOM file
            If tempStr <> "" Then
              ' Ref X Y Rot Layer
              extList = setTextWidth(mod_ref, 15)
              extList = extList + "  " + setTextWidth(mod_val, 25) + "   N/A    N/A "
              extList = extList + "  " + setTextWidth(CStr(mod_x * POSScale), 7)
              extList = extList + "  " + setTextWidth(CStr(-mod_y * POSScale), 7)
              extList = extList + "  " + setTextWidth(CStr(Round(mod_angle, 2)), 6)
              If mod_layer = silkTopLayer Then
                extList = extList + "  Primary Side"
              Else
                extList = extList + "  Secondary Side"
              End If
              Print #posFile, extList
              
              ' 3DBOM File
              Dim tmpScale
              If ext = ".stp" Or ext = ".step" Then
                tmpScale = uni_scale * 2.54
              Else
                tmpScale = uni_scale
              End If
              
              Dim item, key
              key = mod_val + " " + tempStr
              item = dict(key)
              If IsEmpty(item) Then
                item = Array( _
                    mod_ref, _
                    """" + mod_val + """" _
                    + "," + setTextWidth(CStr(scale_x * tmpScale), 8) _
                    + "," + setTextWidth(CStr(scale_y * tmpScale), 8) _
                    + "," + setTextWidth(CStr(scale_z * tmpScale), 8) _
                    + "," + setTextWidth(CStr(ofs_x * BOMScale), 8) _
                    + "," + setTextWidth(CStr(-ofs_y * BOMScale), 8) _
                    + "," + setTextWidth(CStr(ofs_z * BOMScale), 8) _
                    + "," + setTextWidth(CStr(Round(Rot_x, 2)), 8) _
                    + "," + setTextWidth(CStr(Round(Rot_y, 2)), 8) _
                    + "," + setTextWidth(CStr(Round(Rot_z, 2)), 8) _
                    + "," + Utilities.RemoveFileExt(tempStr) _
                  )
              Else
                item(0) = item(0) + "," + mod_ref
              End If
              dict(key) = item
            End If
            
            ' Check if Solidwork model already exist, or need to regenerate it from source
            extList = FindFile(myPath, tempStr + tempStr2, ".sldprt")
            If extList <> "" Then
              If start_time < FileDateTime(extList + ".sldprt") Then
                ext = "USE_SLDPRT"
              ElseIf (ext <> "") And (Not regenerateSLDPRT) Then
                If FileDateTime(Na + ext) <= FileDateTime(extList + ".sldprt") Then
                  ext = "USE_SLDPRT"
                End If
              End If
            End If
            
            Select Case ext
              Case ""
                Na = ""
                
              Case "USE_SLDPRT"
                Na = extList + ".sldprt"
                
              Case ".sldprt"
                Dim nErrors As Long, nWarnings As Long
                Set myPart = swApp.OpenDoc6(Na + ext, swDocPART, _
                  swOpenDocOptions_ReadOnly Or swOpenDocOptions_Silent, _
                  "", nErrors, nWarnings)
                myPart.FeatureManager.InsertScale 0, False, scale_x, scale_y, scale_z
                myPart.ViewZoomtofit2
                Na = Na + tempStr2 + ".sldprt"
                myPart.SaveAs Na
                swApp.CloseDoc GetFileName(myPart.GetPathName())
                Set myPart = Nothing
              
              Case ".wrl"
                Na = Na
                FrmStatus.AppendTODO "Generate Part from " + Na + ext
                
                Set myPart = swApp.NewPart
                'myPart.Visible = False
                GenerateVMRL myPart, Na + ext, scale_x, scale_y, scale_z, genWireFrame
                myPart.ViewZoomtofit2
                Na = Na + tempStr2 + ".sldprt"
                myPart.SaveAs Na
                swApp.CloseDoc GetFileName(myPart.GetPathName())
                Set myPart = Nothing
                
                FrmStatus.PopTODO
                'Part.Visible = PCBCfgForm.PartVisible
              
              Case Else
                ' Check if STEP file exist
                Na = Na
                FrmStatus.AppendTODO "Import part from " + Na + ext
                
                Dim importData As SldWorks.ImportStepData
                Set importData = swApp.GetImportFileData(Na + ext)
                importData.MapConfigurationData = True
                Set myPart = Nothing
                On Error Resume Next
                Set myPart = swApp.LoadFile4(Na + ext, "r", importData, longstatus)
                If Not myPart Is Nothing Then
                  ' Due to process convert STP to WRL file
                  scale_x = 2.54 * scale_x
                  scale_y = 2.54 * scale_y
                  scale_z = 2.54 * scale_z
                  If (scale_x = scale_y) And (scale_x = scale_z) Then
                    If Abs(scale_x - 1) > 0.001 Then
                      myPart.FeatureManager.InsertScale swScaleAboutOrigin, True, scale_x, scale_y, scale_z
                    End If
                  Else
                    myPart.FeatureManager.InsertScale swScaleAboutOrigin, False, scale_x, scale_y, scale_z
                  End If
                  myPart.ViewZoomtofit2
                  Na = Na + tempStr2 + ".sldprt"
                  myPart.SaveAs Na
                  swApp.CloseDoc GetFileName(myPart.GetPathName())
                  Set myPart = Nothing
                  
                  FrmStatus.PopTODO
                Else
                  FrmStatus.PopTODO "FAILED - "
                  Na = ""
                End If
            End Select
            
            ' Generate position and 3dbom files
            If tempStr <> "" Then
              If Na <> "" Then
                tempStr = Na
              Else
                tempStr = FindFile(myPath, tempStr, "")
              End If
            End If
            
            ' Create tranform matrix
            If Na <> "" Then
              compNames.Push Na
              compRef.Push mod_ref
              compVal.Push mod_val
          
              ' Define the transformation matrix. See the IMathTransform API documentation.
              ' Add a rotational diagonal unit matrix (zero rotation) to the transformation matrix
              'mod_angle = 180 - mod_angle
              If mod_layer = silkTopLayer Then
                Rot_x = Rot_x + 180
                Rot_y = -Rot_y
                Rot_z = mod_angle - Rot_z
                mod_angle = 180 - mod_angle
                ofs_x = -ofs_x
              Else
                Rot_z = mod_angle + Rot_z
                mod_angle = -mod_angle
              End If
                            
              Rot_x = Rot_x / 180 * PI
              Rot_y = Rot_y / 180 * PI
              Rot_z = Rot_z / 180 * PI
              mod_angle = mod_angle / 180 * PI
              Cx = Cos(Rot_x)
              Sx = Sin(Rot_x)
              Cy = Cos(Rot_y)
              Sy = Sin(Rot_y)
              Cz = Cos(Rot_z)
              Sz = Sin(Rot_z)
              Ca = Cos(mod_angle)
              Sa = Sin(mod_angle)

              compXforms.Push Cy * Cz
              compXforms.Push -Cy * Sz
              compXforms.Push Sy
              compXforms.Push Sx * Sy * Cz + Cx * Sz
              compXforms.Push Cx * Cz - Sx * Sy * Sz
              compXforms.Push -Sx * Cy
              compXforms.Push Sx * Sz - Cx * Sy * Cz
              compXforms.Push Cx * Sy * Sz + Sx * Cz
              compXforms.Push Cx * Cy
              
              ' Add a translation vector to the transformation matrix
              compXforms.Push ofs_x * Ca - ofs_y * Sa + mod_x
              compXforms.Push ofs_y * Ca + ofs_x * Sa + mod_y
              If mod_layer = silkBottomLayer Then
                compXforms.Push ofs_z + PCB_Thickness / 2 * IncToMeter
              Else
                compXforms.Push -(ofs_z + PCB_Thickness / 2 * IncToMeter)
              End If
              
              ' Add a scaling factor to the transform
              compXforms.Push uni_scale
              
              ' The last three elements of the transformation matrix are unused
              compXforms.Push 0#
              compXforms.Push 0#
              compXforms.Push 0#
            
            End If

          End If
          data_flag = data_flag And 1
          Do While st.Pop() <> "$SHAPE3D"
          Loop
          
        Case Else
          Select Case st.Top()
            Case "$MODULE"
              Select Case a(0)
                Case "Po"
                  If (data_flag And 1) = 0 Then
                    mod_x = CDbl(a(1)) * KiCadScale
                    mod_y = CDbl(a(2)) * KiCadScale
                    mod_angle = CDbl(a(3)) / 10
                    data_flag = data_flag Or 1
                  End If
                
                Case "T0" ' Reference
                  mod_ref = StrJoin(a, 11, UBound(a))
                  mod_ref = (Mid(mod_ref, 2, Len(mod_ref) - 2))
                  mod_layer = CInt(a(9))
                
                Case "T1" ' Value
                  mod_val = StrJoin(a, 11, UBound(a))
                  mod_val = (Mid(mod_val, 2, Len(mod_val) - 2))
                  
                Case "DS"
                  ty2 = CDbl(a(4)) * KiCadScale ' y2
                  tx2 = CDbl(a(3)) * KiCadScale ' x2
                  ty1 = CDbl(a(2)) * KiCadScale ' y1
                  tx1 = CDbl(a(1)) * KiCadScale ' x1
                  If min_x > tx1 Then min_x = tx1
                  If max_x < tx2 Then max_x = tx2
                  If min_y > ty1 Then min_y = ty1
                  If max_y < ty2 Then max_y = ty2
                  
                Case "DC"
                  ty2 = CDbl(a(3)) * KiCadScale ' radius
                  ty1 = CDbl(a(2)) * KiCadScale - ty2
                  tx1 = CDbl(a(1)) * KiCadScale - ty2
                  tx2 = tx1 + 2 * ty2
                  ty2 = ty1 + 2 * ty2
                  If min_x > tx1 Then min_x = tx1
                  If max_x < tx2 Then max_x = tx2
                  If min_y > ty1 Then min_y = ty1
                  If max_y < ty2 Then max_y = ty2
                  
              End Select
            'End Case MODULE
              
            Case "$SHAPE3D"
              Select Case a(0)
                Case "Na"
                  Na = StrJoin(a, 1, UBound(a))
                  Na = (Replace(Mid(Na, 2, Len(Na) - 2), "\\", "\"))
                  data_flag = data_flag Or 2
                
                Case "Sc"
                  scale_x = CDbl(a(1))
                  scale_y = CDbl(a(2))
                  scale_z = CDbl(a(3))
                  
                Case "Of"
                  ofs_x = CDbl(a(1)) * Shape3DScale
                  ofs_y = CDbl(a(2)) * Shape3DScale
                  ofs_z = CDbl(a(3)) * Shape3DScale
                         
                Case "Ro"
                  Rot_x = CDbl(a(1))
                  Rot_y = CDbl(a(2))
                  Rot_z = CDbl(a(3))
              End Select
            'End Case SHAPE3D
            
            Case "$PAD"
              Select Case a(0)
                Case "Sh"
                  tx2 = CDbl(a(3)) * KiCadScale
                  ty2 = CDbl(a(4)) * KiCadScale
                Case "Po"
                  tx1 = CDbl(a(1)) * KiCadScale
                  ty1 = CDbl(a(2)) * KiCadScale
              End Select
              
              
          End Select ' st.Top()
          
      End Select ' s
    End If
  Loop
  FrmStatus.PopTODO
  
  Close #inFile
  Close #posFile
  
  For Each key In dict
    item = dict(key)
    Print #bomFile, """" + item(0) + """," + item(1)
  Next
  Close #bomFile

  If Not compNames.IsEmpty() Then
    ' Add the component to the assembly.
    FrmStatus.setMaxValue compNames.Count
    FrmStatus.setRemaindValue compNames.Count
    vcomponents = Part.AddComponents((compNames.GetArray()), (compXforms.GetArray()))
    FrmStatus.PopTODO
    
    FrmStatus.setMaxValue compNames.Count
    For i = LBound(vcomponents) To UBound(vcomponents)
      FrmStatus.setCurrentValue i
      RelaxForGUI last_time, 0
      Set tmpComp = vcomponents(i)
      tempStr = Replace(compRef.GetValue(i), "@", "at")
      tempStr2 = tempStr + "_" + Replace(compVal.GetValue(i), "@", "at")
      tempStr2 = Replace(tempStr2, "/", ",")
      If RenameComponents Then
        tmpComp.Select2 False, 0
        tmpComp.Name2 = tempStr2
      End If
      tmpComp.ComponentReference = tempStr
    Next i
    FrmStatus.PopTODO
  End If
  
  compNames.Clear
  compXforms.Clear
  compRef.Clear
  compVal.Clear
End Sub

Sub main()
  If PCBCfgForm.PCB_FileName = "" Then
    PCBCfgForm.PCB_FileName = "C:\Documents and Settings\knguyen\Desktop\Projects\IGOR\igor1_prototype\pcb\igor-1_draft0.6 - USB.brd"
    
    KiCadScale = 25.4 / 10000000#
    Shape3DScale = 25.4 / 1000#
    VRMLScale = 25.4 / 10000#
    PCB_Thickness = 0.063
    
    POSScale = 1000000 / 25.4
    BOMScale = 1000 / 25.4
    
    PCBCfgForm.txtPCBThickness = CStr(1000# * PCB_Thickness)
    PCBCfgForm.txtWRLScale = CStr(25.4 / 1000# / VRMLScale)
    PCBCfgForm.txtKiCadScale = CStr(25.4 / 1000# / KiCadScale)
    PCBCfgForm.txtShapeScale = CStr(25.4 / 1000# / Shape3DScale)
    PCBCfgForm.txtPOSScale = CStr(POSScale / 1000 * 25.4)
    PCBCfgForm.txtBOMScale = CStr(BOMScale / 1000 * 25.4)
  End If
  PCBCfgForm.Show
End Sub
