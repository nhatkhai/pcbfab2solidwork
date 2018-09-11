VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack_Dbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'    PCBFab2Solidwork
'    Copyright (C) 2018  NhatKhai L. Nguyen
'
'    Please check LICENSE file for detail.
'
Option Explicit

Private lngTop As Long
Private save_idx As Long
Private myStack() As Double

' to manage increase of array
Private k As Integer
Private Const MAXARRAY = 5

Property Get Count() As Long
  Count = lngTop
End Property

Property Get IsEmpty() As Boolean
  If lngTop = 0 Then IsEmpty = True Else IsEmpty = False
End Property

Function Save()
  save_idx = lngTop
End Function

Function Restore()
  lngTop = save_idx
End Function

Function Push(ByVal value)
  If lngTop >= k * MAXARRAY Then
    k = k + 1
    ReDim Preserve myStack(0 To k * MAXARRAY)
  End If
  myStack(lngTop) = value
  lngTop = lngTop + 1
End Function

Function Pop() As Variant
  If Not IsEmpty() Then
    lngTop = lngTop - 1
    Pop = myStack(lngTop)
  End If
End Function

Function Top() As Variant
  If Not IsEmpty() Then
    Top = myStack(lngTop - 1)
  End If
End Function

Function GetArray() As Variant
  ReDim Preserve myStack(0 To lngTop - 1)
  k = (lngTop - 1) / MAXARRAY
  GetArray = myStack
End Function

Function Clear()
  Erase myStack
  k = 0
  lngTop = 0
End Function

Function GetValue(ByVal idx As Long)
  GetValue = myStack(idx)
End Function

