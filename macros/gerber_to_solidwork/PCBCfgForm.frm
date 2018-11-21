VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PCBCfgForm 
   Caption         =   "KiCad PCB to Solidwork MACRO"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   OleObjectBlob   =   "PCBCfgForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "PCBCfgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'
'    PCBFab2Solidwork
'    Copyright (C) 2018  NhatKhai L. Nguyen
'
'    Please check LICENSE file for detail.
'
Option Explicit

Dim currentPart As Object
Dim swApp As ISldWorks

Private Sub BtnDrillFile_Click()
  Me.DrillFileName = swApp.GetOpenFileName("Select NC Drill file", _
    Me.DrillFileName, "Drill Files (.drl, .ncd)|*.drl;*.ncd|All|*.*|", 0, "", "")
End Sub

Private Sub BtnOutlineFile_Click()
  Me.OutLineFileName = swApp.GetOpenFileName("Select Board outline gerber file", _
    Me.OutLineFileName, "Outline Files|*.g*r;*.pho;*.g*o;*.gm*|All|*.*|", 0, "", "")
End Sub

Private Sub BtnTopSilkFile_Click()
  Me.TopSilkFileName = swApp.GetOpenFileName("Select Top SilkScreen gerber file", _
    Me.TopSilkFileName, "Outline Files|*.g*r;*.pho;*.g*o|All|*.*|", 0, "", "")
End Sub

Private Sub BtnBotSilkFile_Click()
  Me.BotSilkFileName = swApp.GetOpenFileName("Select Bottom SilkScreen gerber file", _
    Me.BotSilkFileName, "Outline Files|*.g*r;*.pho;*.g*o|All|*.*|", 0, "", "")
End Sub

Private Sub BtnBOMFile_Click()
  Me.BOMFileName = swApp.GetOpenFileName("Select 3D BOM File", _
    Me.BOMFileName, "3D BOM Files|*.csv;*.bom|All|*.*|", 0, "", "")
End Sub

Private Sub BtnPosFile_Click()
  Me.PosFileName = swApp.GetOpenFileName("Select Position File", _
    Me.PosFileName, "Position Files|*.csv;*.pos;*.xyr|All|*.*|", 0, "", "")
End Sub

Private Sub cbScaleStyle_Change()
  If LCase(cbScaleStyle.Text) = "kicad" Then
    txtDrillScale.Text = "1" ' 2.4 Format (unit/in)
    txtGerbScale.Text = "1"  ' 2.4 Format (unit/in)
    txtPosScale.Text = "1"       ' unit/in
    txtPosAngleScale.Text = "-1" ' unit/degree
    txtWRLScale.Text = "10"      ' unit/in
    Gerber_To_3D.POS_RefColIdx = 0
    Gerber_To_3D.POS_PosXColIdx = 2
    Gerber_To_3D.POS_PosYColIdx = 3
    Gerber_To_3D.POS_RotColIdx = 4
    Gerber_To_3D.POS_SideColIdx = 5
  ElseIf LCase(cbScaleStyle.Text) = "pad" Then
    txtDrillScale.Text = "1" ' 2.4 Format (unit/in)
    txtGerbScale.Text = "1"  ' 2.4 Format (unit/in)
    txtPosScale.Text = "1000"    ' unit/in
    txtPosAngleScale.Text = "1"  ' unit/degree
    txtWRLScale.Text = "1"       ' unit/in
  End If
End Sub

Private Sub PartVisible_Click()
  If Not currentPart Is Nothing Then
    currentPart.Visible = Me.PartVisible
  End If
End Sub

Private Sub run_Click()
  Dim boardFilename As String
  Dim Part As IAssemblyDoc
  Dim myPart As IPartDoc
  Dim t1, t2
  Dim doAssembly As Boolean
  Dim old_userSketchInference As Boolean
  Dim old_userExtRefUpdateCompNames As Boolean
  
  If Len(Me.DrillFileName.Text) = 0 And Len(Me.OutLineFileName) = 0 Then
    MsgBox ("At least drill file, or board out file need to be specified")
    Exit Sub
  End If
  
  FrmStatus.Show

  ' Make sure only one instance running at the time
  If Me.run.Caption <> "Run" Then Exit Sub
  Me.run.Caption = "Runing..."
  
  FrmStatus.Reset
  Set currentPart = Nothing
  
  If Len(Me.DrillFileName.Text) <> 0 Then
    ' Check time stamp and file exist
    boardFilename = RemoveFileExt(Me.DrillFileName)
    On Error GoTo DrillFileNotExist
    t1 = FileDateTime(Me.DrillFileName)
  ElseIf Len(Me.OutLineFileName.Text) <> 0 Then
    boardFilename = RemoveFileExt(Me.OutLineFileName)
    On Error GoTo OutLineFileNotExist
    t1 = FileDateTime(Me.OutLineFileName)
  End If

  t2 = t1
  If Not Me.AlwaysGenPCBPart Then
    On Error Resume Next
    t2 = FileDateTime(boardFilename + ".sldprt")
  End If
  On Error GoTo 0
  
  old_userSketchInference = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swSketchInference)
  old_userExtRefUpdateCompNames = swApp.GetUserPreferenceToggle(swUserPreferenceToggle_e.swExtRefUpdateCompNames)
  swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swExtRefUpdateCompNames, False
  swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchInference, False
  
  ' Generate TODO list to show user
  If t2 <= t1 Then
    FrmStatus.PushTODO "Generate PCB Part " + boardFilename + ".sldprt"
  End If
  If Me.BOMFileName <> "" And Me.PosFileName <> "" Then
    FrmStatus.PushTODO "Generate Assembly " + boardFilename + ".sldasm"
    doAssembly = True
  End If
  
  ' Generate PCB Part
  If t2 <= t1 Then
    Set myPart = swApp.NewPart
    myPart.Visible = Me.PartVisible
    Set currentPart = myPart
    If Not (myPart Is Nothing) Then
      GeneratePCB myPart, Me.DrillFileName, Me.OutLineFileName, _
        minHole:=CDbl(txtMinHole.Text)
      
      If Me.TopSilkFileName <> "" Then _
        GenerateSilk myPart, Me.TopSilkFileName _
            , (0.001 + Gerber_To_3D.PCB_Thickness / 2) * InchToSW _
            , "TopSilkScreen"
      
      If Me.BotSilkFileName <> "" Then _
        GenerateSilk myPart, Me.BotSilkFileName _
          , -(0.001 + Gerber_To_3D.PCB_Thickness / 2) * InchToSW _
          , "BottomSilkScreen"
      
      myPart.Visible = True
      myPart.ViewZoomtofit2
      myPart.SaveAs boardFilename + ".sldprt"
      
      Set currentPart = Nothing
      If doAssembly Then
        swApp.CloseDoc GetFileName(myPart.GetPathName())
        Set myPart = Nothing
      End If
    Else
      MsgBox ("Can't create new part")
    End If
    FrmStatus.PopTODO
    
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swExtRefUpdateCompNames, old_userExtRefUpdateCompNames
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchInference, old_userSketchInference
    
    
    ' Generate Assembly on PCB Board
    If doAssembly Then
      Set Part = swApp.NewAssembly
      Part.Visible = Me.PartVisible
      Set currentPart = Part
      If Not (Part Is Nothing) Then
        GeneratePCBAssembly Part, boardFilename, _
          Me.PosFileName, Me.BOMFileName, _
          Me.overwriteVRML, _
          Me.genMinMaxBox, _
          False, _
          Me.RenameComponents
        Part.Visible = True
        Part.ViewZoomtofit2
        Part.SaveAs boardFilename + ".sldasm"
        Set currentPart = Nothing
        'swApp.CloseDoc GetFileName(Part.GetPathName())
        'Set Part = Nothing
      Else
        MsgBox ("Can't create new assembly")
      End If
      FrmStatus.PopTODO
    End If
    
  End If
  Me.Hide
  FrmStatus.PushTODO "Done"
  Me.run.Caption = "Run"
  Exit Sub

DrillFileNotExist:
  Me.run.Caption = "Run"
  MsgBox ("Drill file """ + Me.DrillFileName + """ not found!")
  Me.run.Caption = "Run"
  FrmStatus.PushTODO "Done"

OutLineFileNotExist:
  Me.run.Caption = "Run"
  MsgBox ("Outline file """ + Me.OutLineFileName + """ not found!")
  Me.run.Caption = "Run"
  FrmStatus.PushTODO "Done"
End Sub

Private Sub txtDrillScale_Change()
  Dim n
  On Error Resume Next
  n = 1 / CDbl(Me.txtDrillScale)
  Gerber_To_3D.DrillScale = n
  Me.txtDrillScale = CStr(1 / Gerber_To_3D.DrillScale)
End Sub

Private Sub txtGerbScale_Change()
  Dim n
  On Error Resume Next
  n = 1 / CDbl(Me.txtGerbScale)
  Gerber_To_3D.GerbScale = n
  Me.txtGerbScale = CStr(1 / Gerber_To_3D.GerbScale)
End Sub

Private Sub txtMinHole_Change()
  Dim n
  On Error Resume Next
  n = 0.01
  n = CDbl(Me.txtMinHole)
  Me.txtMinHole = CStr(n)
End Sub

Private Sub txtPCBThickness_Change()
  Dim n
  On Error Resume Next
  n = CDbl(Me.txtPCBThickness) / 1000
  Gerber_To_3D.PCB_Thickness = n
  Me.txtPCBThickness = CStr(Gerber_To_3D.PCB_Thickness * 1000)
End Sub

Private Sub txtPosAngleScale_Change()
  Dim n
  On Error Resume Next
  n = CDbl(Me.txtPosAngleScale)
  Gerber_To_3D.AngScale = n
  Me.txtPosAngleScale = CStr(Gerber_To_3D.AngScale)
End Sub

Private Sub txtPosScale_Change()
  Dim n
  On Error Resume Next
  n = Gerber_To_3D.InchToSW / CDbl(Me.txtPosScale)
  Gerber_To_3D.POSScale = n
  Me.txtPosScale = CStr(Gerber_To_3D.InchToSW / Gerber_To_3D.POSScale)
End Sub

Private Sub txtWRLScale_Change()
  Dim n
  On Error Resume Next
  n = Gerber_To_3D.InchToSW / CDbl(Me.txtWRLScale)
  Gerber_To_3D.VRMLScale = n
  Me.txtWRLScale = CStr(Gerber_To_3D.InchToSW / Gerber_To_3D.VRMLScale)
End Sub


Private Sub UserForm_Initialize()
  Set swApp = Application.SldWorks
  cbScaleStyle.AddItem ("PAD")
  cbScaleStyle.AddItem ("KiCad")
  cbScaleStyle.Text = "KiCad"
  cbScaleStyle_Change
  
  If swApp Is Nothing Then
    err.Raise 1000, "Initialize Error", "Solidword Application not found"
  End If
End Sub
