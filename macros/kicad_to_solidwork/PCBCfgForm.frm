VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PCBCfgForm 
   Caption         =   "KiCad PCB to Solidwork MACRO"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
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

Private Sub BtnPCBFile_Click()
  Set swApp = Application.SldWorks
  If swApp Is Nothing Then
    MsgBox "Solidworks Application not found"
    Exit Sub
  End If
    
  Me.PCB_FileName = swApp.GetOpenFileName("Select Layout file", Me.PCB_FileName, "KiCad Layout (*.brd)|*.brd", 0, "", "")
End Sub



Private Sub PartVisible_Click()
  If Not currentPart Is Nothing Then
    currentPart.Visible = Me.PartVisible
  End If
End Sub


Private Sub run_Click()
  Set swApp = Application.SldWorks
  If swApp Is Nothing Then
    MsgBox "Solidworks Application not found"
    Exit Sub
  End If
  
  FrmStatus.Show
  FrmStatus.Reset
  Set currentPart = Nothing
  If Me.run.Caption <> "Run" Then Exit Sub
  
  If Not IsNumeric(Me.txtCmpSize) Then
    MsgBox ("Min Size must be a real number in mil^2")
    Exit Sub
  End If
  
  If Not IsNumeric(Me.txtCmpSizeAssembly) Then
    MsgBox ("Min Size must be a real number in mil^2")
    Exit Sub
  End If
  
  If Me.PCB_FileName = "" Then
    Me.PCB_FileName = swApp.GetOpenFileName("Select Layout file", Me.PCB_FileName, "KiCad Layout (*.brd)|*.brd", 0, "", "")
  End If
    
  Me.run.Caption = "Runing..."
  If Me.PCB_FileName <> "" Then
    
    Dim partFile As String
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swExtRefUpdateCompNames, False
    swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchInference, False
    
    Dim Part As IAssemblyDoc
    Dim myPart As IPartDoc
    Dim t1, t2
    
    ' Check time stamp and file exist
    partFile = RemoveFileExt(Me.PCB_FileName) + ".sldprt"
    On Error GoTo FileNotExist
    t1 = FileDateTime(Me.PCB_FileName)
    t2 = t1
    If Not Me.AlwaysGenPCBPart Then
      On Error Resume Next
      t2 = FileDateTime(partFile)
    End If
    On Error GoTo 0
    
    If t2 <= t1 Then FrmStatus.PushTODO "Generate PCB Part"
    If Me.CheckBoxAssembly Then FrmStatus.PushTODO "Generate Assembly"
    
    ' Generate PCB Part
    If t2 <= t1 Then
      Set myPart = swApp.NewPart
      myPart.Visible = Me.PartVisible
      Set currentPart = myPart
      If Not (myPart Is Nothing) Then
        GeneratePCB myPart, Me.PCB_FileName, _
          Me.GenSilks, Me.SilksDSSketchOnly, _
          Me.GenMinMaxSilks, Me.genText, _
          CDbl(Me.txtCmpSize) / 1000
        myPart.Visible = True
        myPart.ViewZoomtofit2
        myPart.SaveAs partFile
        Set currentPart = Nothing
        If Me.CheckBoxAssembly Then
          swApp.CloseDoc GetFileName(myPart.GetPathName())
          Set myPart = Nothing
        End If
      Else
        MsgBox ("Can't create new part")
      End If
      FrmStatus.PopTODO
    End If
    
    ' Generate Assembly on PCB Board
    If Me.CheckBoxAssembly Then
      Set Part = swApp.NewAssembly
      Part.Visible = Me.PartVisible
      Set currentPart = Part
      If Not (Part Is Nothing) Then
        GeneratePCBAssembly Part, Me.PCB_FileName, _
          Me.forceRecreateSLDPRT, _
          Me.genMinMaxBox, Me.genWireFrame, _
          CDbl(Me.txtCmpSizeAssembly) / 1000, _
          Me.RenameComponents
        partFile = RemoveFileExt(Me.PCB_FileName) + ".sldasm"
        Part.Visible = True
        Part.ViewZoomtofit2
        Part.SaveAs partFile
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
  Me.run.Caption = "Run"
  FrmStatus.PushTODO "Done"
  Exit Sub

FileNotExist:
  Me.run.Caption = "Run"
  MsgBox ("file """ + Me.PCB_FileName + """ not found!")
End Sub

Private Sub txtBOMScale_Change()
  Dim n
  On Error Resume Next
  n = CDbl(Me.txtBOMScale) * 1000 / 25.4
  KiCad_to_SolidWork.BOMScale = n
  Me.txtBOMScale = CStr(KiCad_to_SolidWork.BOMScale / 1000 * 25.4)
End Sub

Private Sub txtKiCadScale_Change()
  Dim n
  On Error Resume Next
  n = 25.4 / 1000# / CDbl(Me.txtKiCadScale)
  KiCad_to_SolidWork.KiCadScale = n
  Me.txtKiCadScale = CStr(25.4 / 1000# / KiCad_to_SolidWork.KiCadScale)
End Sub

Private Sub txtPCBThickness_Change()
  Dim n
  On Error Resume Next
  n = CDbl(Me.txtPCBThickness) / 1000#
  KiCad_to_SolidWork.PCB_Thickness = n
  Me.txtPCBThickness = CStr(KiCad_to_SolidWork.PCB_Thickness * 1000#)
End Sub

Private Sub txtPOSScale_Change()
  Dim n
  On Error Resume Next
  n = CDbl(Me.txtPOSScale) * 1000 / 25.4
  KiCad_to_SolidWork.POSScale = n
  Me.txtPOSScale = CStr(KiCad_to_SolidWork.POSScale / 1000 * 25.4)

End Sub

Private Sub txtShapeScale_Change()
  Dim n
  On Error Resume Next
  n = 25.4 / 1000# / CDbl(Me.txtShapeScale)
  KiCad_to_SolidWork.Shape3DScale = n
  Me.txtShapeScale = CStr(25.4 / 1000# / KiCad_to_SolidWork.Shape3DScale)
End Sub

Private Sub txtWRLScale_Change()
  Dim n
  On Error Resume Next
  n = 25.4 / 1000# / CDbl(Me.txtWRLScale)
  KiCad_to_SolidWork.VRMLScale = n
  Me.txtWRLScale = CStr(25.4 / 1000# / KiCad_to_SolidWork.VRMLScale)
End Sub
