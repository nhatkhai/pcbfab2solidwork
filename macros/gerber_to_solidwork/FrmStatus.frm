VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmStatus 
   Caption         =   "Status"
   ClientHeight    =   4704
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   OleObjectBlob   =   "FrmStatus.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmStatus"
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
Dim todoCount As Long
Dim mymax As Double
Dim myvalue As Double

Public Sub Reset()
  todoCount = 0
  Me.TODOList.Clear
  Me.DONEList.Clear
  Me.ProgressText = ""
  Me.ProgressBar.Caption = ""
End Sub

Public Sub AppendTODO(s As String)
  Me.TODOList.AddItem s, todoCount
  todoCount = todoCount + 1

  Me.ProgressText = s
  Me.Repaint
End Sub

Public Sub PushTODO(s As String)
  Me.TODOList.AddItem s, 0
  
  Me.ProgressText = Me.TODOList.List(todoCount)
  todoCount = todoCount + 1
  Me.Repaint
End Sub

Public Sub PopTODO()
  Dim s As String
  
  If (todoCount > 0) Then
    todoCount = todoCount - 1
    s = Me.TODOList.List(todoCount)
    Me.TODOList.RemoveItem todoCount
    Me.DONEList.AddItem s, 0
  End If

  If (todoCount > 0) Then
    Me.ProgressText = Me.TODOList.List(todoCount - 1)
  Else
    Me.ProgressText = ""
  End If
  Me.DONEList.ListIndex = 0
  Me.Repaint
End Sub

Public Sub setCurrentValue(ByVal value As Long, Optional ByVal max As Long = -1)
  If max > value Then mymax = max
  myvalue = value
  updateProgressBar
  Me.Repaint
End Sub

Public Sub setMaxValue(ByVal max As Long)
  mymax = max
  myvalue = 0
  updateProgressBar
  Me.Repaint
End Sub

Public Sub setRemaindValue(ByVal value As Long)
  myvalue = mymax - value
  updateProgressBar
  Me.Repaint
End Sub

Public Sub IncValue()
  myvalue = myvalue + 1
  updateProgressBar
  Me.Repaint
End Sub

Sub updateProgressBar()
  Dim i As Integer
  Dim v As Integer
  Dim s As String
  
  If mymax <= 0 Then
    Me.ProgressBar.Caption = ""
    Me.ProgressBar2.Caption = ""
    Exit Sub
  End If
  
  v = Round(myvalue / mymax * 100, 0)
  Me.ProgressBar.Caption = Str(myvalue) + "/" + Str(mymax) + " (" + Str(v) + "%)"
  
  v = v * 30 / 100
  For i = 1 To 30
    If i < v Then
      s = s + "+"
    Else
      s = s + "-"
    End If
  Next i
  Me.ProgressBar2.Caption = "[" + s + "]"
End Sub
