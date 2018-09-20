Attribute VB_Name = "Utilities"
'
'    PCBFab2Solidwork
'    Copyright (C) 2018  NhatKhai L. Nguyen
'
'    Please check LICENSE file for detail.
'
Option Explicit

Public Const PI = 3.14159265358979
Public Const SearchPaths = "c:\Program Files\KiCad\share\modules\packages3d"


Function StrJoin(a() As String _
  , fromIdx As Integer _
  , toIdx As Integer _
  , Optional joinStr As String = " " ) As String

  Dim i As Integer
  Dim s As String
  s = a(fromIdx)
  For i = fromIdx + 1 To toIdx
    s = s + joinStr + a(i)
  Next i
  StrJoin = s
End Function


Function FindFile(Path As String, FileName As String _
  , ByVal ext As String) As String

  Dim tmpStr, p, rs
  FindFile = ""
  
  On Error GoTo Error1
  tmpStr = FileName + ext
  
  If (Right(Path, 1) <> "\") Then Path = Path + "\"
  
  p = Dir$(Path + tmpStr)
  If (p <> "") Then
    FindFile = Path + FileName
    Exit Function
  End If
  
  p = Dir$(tmpStr)
  If (p <> "") Then
    FindFile = FileName
    Exit Function
  End If

  For Each rs In Split(SearchPaths, ";")
    If (Right(rs, 1) <> "\") Then rs = rs + "\"
    p = Dir$(rs + tmpStr)
    If p <> "" Then
      FindFile = rs + FileName
      Exit Function
    End If
  Next rs
  Exit Function
Error1:
  p = ""
  Resume Next
End Function


Function RemoveFileExt(ByVal strPath As String) As String
    On Error Resume Next
    RemoveFileExt = strPath
    RemoveFileExt = Left$(strPath, InStrRev(strPath, ".") - 1)
End Function


Function GetFileExt(ByVal strPath As String) As String
    On Error Resume Next
    GetFileExt = ""
    GetFileExt = Mid$(strPath, InStrRev(strPath, "."))
End Function


Function FileNameNoExt(strPath As String) As String
    On Error Resume Next
    FileNameNoExt = strPath
    Dim strTemp As String
    strTemp = Mid$(strPath, InStrRev(strPath, "\") + 1)
    FileNameNoExt = Left$(strTemp, InStrRev(strTemp, ".") - 1)
End Function
 

Function GetFileName(strPath As String) As String
    On Error Resume Next
    GetFileName = strPath
    GetFileName = Mid$(strPath, InStrRev(strPath, "\") + 1)
End Function
 

Function FilePath(strPath As String) As String
    On Error Resume Next
    FilePath = Left$(strPath, InStrRev(strPath, "\"))
End Function


Function ReadCSVRow(ByVal row As String) As Variant
  Dim tmpa, tmpb, tmpj, tmpi, sz, tmps
  Dim cols()
  
  sz = 0
  tmpa = Split(row, """")
  
  For tmpi = 0 To UBound(tmpa)
    If tmpi Mod 2 = 0 Then
      tmps = Trim(tmpa(tmpi))
      
      If tmps <> "" Then
        If (tmpi < UBound(tmpa)) And (Right(tmps, 1) = ",") Then tmps = Left(tmps, Len(tmps) - 1)
        If (tmpi > 0) And (Left(tmps, 1) = ",") Then tmps = Mid(tmps, 2)
      
        tmpb = Split(tmps, ",")
        ReDim Preserve cols(sz + UBound(tmpb))
        For tmpj = 0 To UBound(tmpb)
          cols(tmpj + sz) = tmpb(tmpj)
        Next
        sz = sz + tmpj
      End If
    Else
      ReDim Preserve cols(sz)
      cols(sz) = tmpa(tmpi)
      sz = sz + 1
    End If
  Next
  
  ReadCSVRow = cols
End Function


Function ReadSpaceSepVecRow(ByVal row As String) As Variant
  Dim tmpa, tmpb, tmpj, tmpi, sz, tmps
  Dim cols()
  
  sz = 0
  tmpa = Split(row, """")
  
  For tmpi = 0 To UBound(tmpa)
    If tmpi Mod 2 = 0 Then
      tmps = Trim(tmpa(tmpi))
      
      If tmps <> "" Then
        tmpb = Split(tmps, " ")
        ReDim Preserve cols(sz + UBound(tmpb))
        For tmpj = 0 To UBound(tmpb)
          tmps = Trim(tmpb(tmpj))
          If tmps <> "" Then
            cols(sz) = tmps
            sz = sz + 1
          End If
        Next
        If sz > 0 Then ReDim Preserve cols(sz - 1)
      End If
    Else
      ReDim Preserve cols(sz)
      cols(sz) = tmpa(tmpi)
      sz = sz + 1
    End If
  Next
  
  ReadSpaceSepVecRow = cols
End Function


Function Read3DBOMFile(FileName As String) As Object
  Dim dict
  Dim s, mypath, ref
  Dim cols(), refs
  Dim lst As Stack_Objects
  Dim info3d(9)
  Dim inFile, line As Integer
  Dim maxCol
  
  'Const RefColIdx = 2
  'Const ScaleColIdx = 9
  'Const OfsColIdx = 12
  'Const RotColIdx = 15
  'Const ModleFileColIdx = 18

  maxCol = Gerber_To_3D.BOM_ModleFileColIdx
  If maxCol < Gerber_To_3D.BOM_RotColIdx + 3 Then maxCol = Gerber_To_3D.BOM_RotColIdx + 3
  If maxCol < Gerber_To_3D.BOM_OfsColIdx + 3 Then maxCol = Gerber_To_3D.BOM_OfsColIdx + 3
  If maxCol < Gerber_To_3D.BOM_ScaleColIdx + 3 Then maxCol = Gerber_To_3D.BOM_ScaleColIdx + 3
  If maxCol < Gerber_To_3D.BOM_RefColIdx Then maxCol = Gerber_To_3D.BOM_RefColIdx
  
  Set dict = CreateObject("Scripting.Dictionary")
  
  inFile = FreeFile
  Open FileName For Input As #inFile
  mypath = FilePath(FileName)
  line = 0
  Do While Not EOF(inFile)
    Line Input #inFile, s
    line = line + 1
    
    cols = ReadCSVRow(s)
    If UBound(cols) >= maxCol Then
      refs = Split(cols(Gerber_To_3D.BOM_RefColIdx), ",")
      For Each ref In refs
        ref = UCase(Trim(ref))
        If ref <> "" Then
          info3d(0) = cols(Gerber_To_3D.BOM_ScaleColIdx)     ' Scale X
          info3d(1) = cols(Gerber_To_3D.BOM_ScaleColIdx + 1) ' Scale Y
          info3d(2) = cols(Gerber_To_3D.BOM_ScaleColIdx + 2) ' Scale Z
          info3d(3) = cols(Gerber_To_3D.BOM_OfsColIdx)       ' Ofs X
          info3d(4) = cols(Gerber_To_3D.BOM_OfsColIdx + 1)   ' Ofs Y
          info3d(5) = cols(Gerber_To_3D.BOM_OfsColIdx + 2)   ' Ofs Z
          info3d(6) = cols(Gerber_To_3D.BOM_RotColIdx)       ' Rot X
          info3d(7) = cols(Gerber_To_3D.BOM_RotColIdx + 1)   ' Rot Y
          info3d(8) = cols(Gerber_To_3D.BOM_RotColIdx + 2)   ' Rot Z
          info3d(9) = cols(Gerber_To_3D.BOM_ModleFileColIdx) ' 3D Model files (STEP, SLDPRT or VMRL)
          If info3d(9) <> "" Then
            If dict.Exists(ref) Then
              Set lst = dict.Item(ref)
            Else
              Set lst = New Stack_Objects
              dict.Add ref, lst
            End If
            lst.Push info3d
          End If
        End If
      Next
    End If
  Loop
  Close #inFile
  
  Set Read3DBOMFile = dict
End Function


' This help script run faster by slow down Solidwork update rate while 
' generating many Solidwork Objects.
Function RelaxForGUI(ByRef last_time As Long, ByVal interval As Long) As Boolean
  If Timer >= last_time Then
    DoEvents
    last_time = Timer + interval
    RelaxForGUI = True
  Else
    RelaxForGUI = False
  End If
End Function


' Extract a number in strNum from offset of StartIdx in the string.
'
' The number will be extracted base on Gerber FS format. It then apply the
' UnitScale to convert into the usable Unit.
'
' @example:
'   * GerberNumber("%ADD10C", StartIdx:Idx) --> 12, and Idx change from 5 to 7
'   * GerberNumber( "X123Y-100D02*", 25.4
'                 , 1e-4, 6  <-- 2.4 FS Format yeild 1e-4, 6
'                                4.6 FS Format yeild 1e-6, 10
'                                2.5 FS Format yeild 1e-6, 7 
'                 , False
'                 , Idx) -> 0.0123*25.4, and Idx change from 2 to 5
' 
' @param strNum [in] a string containt gerber number that need to be
'           extracted
' @param UnitScale [in] a convertion factor for convert extracted number
'           into usable unit. Default is 1
' @param CorrectScale [in] Specify adjust scale. This value is 
'           1/10^(#DECIMAL_DIGIT). Where #DECIMAL_DIGIT is
'           specified in Gerber FS command. Default is 1
' @param NumDigit [in] Specify maximum total number of digit can be.
'           This value is #DECIMAL_DIGIT + #INTEGER_DIGIT. There are
'           specified in Gerber FS command. Default is 6
' @param Leading [in] Specify where Gerber number in leading zeros (True) or
'           trailling zeros (False) format. Default is True
' @param StartIdx [in/out] Specify starting offset in strNum for extract
'           the number. StartIdx will be advanced to the position after the
'           extracted number in the string.
'
' @return The extract gerber number in specified unit base on UnitScale
'
Function GerberNumber(ByVal strNum As String _
  , Optional ByVal UnitScale As Double = 1 _
  , Optional ByVal CorrectScale As Double = 1 _
  , Optional ByVal NumDigit As Integer = 6 _
  , Optional ByVal Leading As Boolean = False _
  , Optional ByRef StartIdx = 1) As Double
  Dim num As Double
  Dim i As Integer
  Dim dot As Boolean
  
  dot = False
  i = StartIdx

  Select Case Mid(strNum, i, 1)
    Case "-", "+"
      i = i + 1
  End Select

  Do
    Select Case Mid(strNum, i, 1)
      Case "0" To "9"
      Case "."
        If dot Then Exit Do
        dot = True
      Case Else
        Exit Do
    End Select
    i = i + 1
  Loop Until i > Len(strNum)
  
  If (i = StartIdx) Then
    num = 0
  Else
    If Leading And (Not Dot) Then
      num = CDbl(Left( Mid(strNum, StartIdx, i - StartIdx) _
                     + String(NumDigit, "0") _
                     , NumDigit))
    Else
      num = CDbl(Mid(strNum, StartIdx, i - StartIdx))
    End If
  End If
  
  StartIdx = i
  
  If InStr(strNum, ".") = 0 Then
    num = num * CorrectScale
  End If

  GerberNumber = num * UnitScale
  
End Function


' Function extract Gerber CMD string from strNum starting at position
' StartIdx in the string.
'
' @example:
'   * GerberCMD("%ADD10C,0.1*%", Idx) -> "%ADD", and Idx change from 1 to 5
'   * GerberCMD("%ADD10C,0.1*%", Idx) -> "C", and Idx change from 7 to 9
'   * GerberCMD("G70*", Idx, "G" ) -> "G", and Idx change from 1 to 2
'   * GerberCMD("G70*", Idx, Null) -> "G", and Idx change from 1 to 2
'   * GerberCMD("G70*", Idx, "E" ) -> Null, and Idx stay at 1
'
' @param strNum [in] a line in Gerber file
' @param StartIdx [in/out] Specify initial location for extract Gerber CMD.
'           StartIdx will be advanced after the extract command string.
' @param testCMD [in] A string use to test against the extracted Gerber
'           CMD. If Null, test will be omitted.
'
' @return Null when the test faild, the StartIdx value will not be
'   modified. Return a extract Gerber CMD string when the test is matched
'   or omitted
'
Function GerberCMD(ByVal strNum As String _
  , Optional ByRef StartIdx = 1 _
  , Optional ByRef testCMD = Null)
  Dim i As Integer, j As Integer
  
  Do While Mid(strNum, StartIdx, 1) = ","
    StartIdx = StartIdx + 1
  Loop
  
  i = StartIdx
  Do
    Select Case Mid(strNum, StartIdx, 1)
      Case "-", "+", ".", "0" To "9"
        j = StartIdx
        Exit Do
      Case ","
        j = StartIdx
        StartIdx = StartIdx + 1
        Exit Do
    End Select
    
    StartIdx = StartIdx + 1
    If StartIdx > Len(strNum) Then
       j = StartIdx
      Exit Do
    End If
  Loop
  
  GerberCMD = Mid(strNum, i, j - i)
  If StrComp(GerberCMD, testCMD) <> 0 Then
    StartIdx = i
    GerberCMD = Null
  End If
  
End Function


' Function return angle in radiant between [-pi;pi] with given x and y
' coordinate, using following scheme:
' 
' x<0, y>0:  angle = 180+atn(y/x)  =  90-atn(x/y)
' x>0, y>0:  angle =     atn(y/x)  =  90-atn(x/y)
' x>0, y<0:  angle =     atn(y/x)  = -90-atn(x/y)
' x<0, y<0:  angle =-180+atn(y/x)) = -90-atn(x/y)
'
'                   |
'                   |
'        90-atn(x/y)|  90-atn(x/y)
'       180+atn(y/x)|  atn(y/x)
'   ----------------+-----------------
'      -180+atn(y/x)|  atn(y/x)
'      -90 -atn(x/y)| -90-atn(x/y)
'                   |
'                   |
'
Function angle(x As Double, y As Double) As Double
  If y >= 0 Then
    If x >= 0 Then
      ' First Quadrant
      If x > y Then
        angle = Atn(y / x)
      ElseIf y = 0 Then
        angle = 0
      Else
        angle = PI / 2 - Atn(x / y)
      End If
    Else
      ' Second Quadrant
      If (-x > y) Then
        angle = PI + Atn(y / x)
      Else
        angle = PI / 2 - Atn(x / y)
      End If
    End If
  Else
    If x >= 0 Then
      ' Fourth Quadrant
      If (x > -y) Then
        angle = Atn(y / x)
      Else
        angle = -PI / 2 - Atn(x / y)
      End If
    Else
      ' Third Quadrant
      If (-x > -y) Then
        angle = -PI + Atn(y / x)
      Else
        angle = -PI / 2 - Atn(x / y)
      End If
    End If
  End If
End Function


' Normalize angle x to radiant value between [-pi;pi]
'
Function normalizeAngle(ByVal x As Double)
  Do While x > PI
    x = x - 2 * PI
  Loop
  Do While x < -PI
    x = x + 2 * PI
  Loop
  normalizeAngle = x
End Function


Sub Rotate2D(ByRef x As Double, ByRef y As Double, ByVal angle As Double, ByVal OFS_x As Double, ByVal OFS_y As Double)
  Dim ts, tc, tx, ty
  ts = Sin(angle * PI / 180)
  tc = Cos(angle * PI / 180)
  tx = x * tc + y * ts + OFS_x
  ty = y * tc - x * ts + OFS_y
  x = tx
  y = ty
End Sub


Function SingleQuadrantArcCenter(ByVal x1 As Double, ByVal y1 As Double, _
  ByRef dcx As Double, ByRef dcy As Double, _
  ByVal x2 As Double, ByVal y2 As Double, _
  clockwise As Boolean)
  
  Dim ac(0 To 3) As Double
  Dim aend As Double
  Dim i As Integer
  
  aend = angle(x2 - x1, y2 - y1)
  ac(0) = normalizeAngle(angle(dcx, dcy) - aend)
  ac(1) = normalizeAngle(angle(-dcx, dcy) - aend)
  ac(3) = normalizeAngle(angle(-dcx, -dcy) - aend)
  ac(2) = normalizeAngle(angle(dcx, -dcy) - aend)
  
  If Not clockwise Then
    For i = 0 To 3
      If ac(i) >= PI / 4 And ac(i) < PI / 2 Then
        Exit For
      End If
    Next i
  Else
    For i = 0 To 3
      If ac(i) <= -PI / 4 And ac(i) > -PI / 2 Then
        Exit For
      End If
    Next i
  End If
    
  If i = 1 Or i = 3 Then dcx = -dcx
  If i = 2 Or i = 3 Then dcy = -dcy
End Function


' This is just manual one op code testing
Function REMOVE_TEST()
  'Dim d
  'Set d = Read3DBOMFile("C:\Documents and Settings\knguyen\Desktop\Projects\IGOR\igor1_prototype\pcb\Gerber_MAIN_PCB\Rev3\REMOVE_BOM.csv")
  Dim dict
  Dim ob1 As Stack_Objects
  Dim ob2 As Stack_Objects
  Dim ob3
  Dim lst(1 To 4)
  
  Set dict = CreateObject("Scripting.Dictionary")
  Set ob1 = New Stack_Objects
  Set ob2 = New Stack_Objects
  dict.Add "1", ob1
  dict.Add "2", ob2
  ob1.Push "Khai1"
  ob1.Push "Khai2"
  For Each ob3 In dict.Item("0").GetArray()
    ob2.Push ob3
  Next ob3
  
  Dim r, ss As String, a As Integer
  a = 1
  ss = "M70,LZ"
  r = GerberCMD(ss, a)
  r = GerberNumber(ss, 1, 1, , True, a)
  r = GerberCMD(ss, a, "TZ")
  r = GerberCMD(ss, a, "LZ")
End Function

