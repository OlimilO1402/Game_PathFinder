VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Landscape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '2007_06_14 Zeilen: 199
Public Enum LandscapeElement
  Mountain = &H1&   '0000 0001 'Grau       'Felsengebirge
  Valley = &H2&     '0000 0010 'Dunkelgrün 'Wald und Wiese
  InitPoint = &H6&  '0000 0110 'Blue
  TrackBad = &HA&   '0000 1010 'Rot        '
  TrackFind = &H12& '0001 0010 'Gelb
  TrackGood = &H22& '0010 0010 'Grün
  GoalPoint = &H42& '0100 0010 'White'Black
  GoalFound = &H82& '1000 0010 'Blinking
End Enum
Private mLandArr() As LandscapeElement 'Breite * Höhe
Public Event Draw()
Public Event DrawBlock(X As Long, Y As Long)

Public Sub NewC(ByVal aWidth As Long, ByVal aHeight As Long)
  Call SetArea(aWidth, aHeight)
  Call GenerateRndLand
End Sub
Public Sub SetArea(aWidth As Long, aHeight As Long)
  ReDim mLandArr(0 To aWidth - 1, 0 To aHeight - 1)
End Sub
Public Sub GenerateRndLand(Optional ByVal ProzValley As Double = 0.67)
Dim i As Long, j As Long
Dim d As Double, p As Double
  If ProzValley > 1 Then p = ProzValley / 100 Else p = ProzValley
  For i = 0 To Width - 1
    For j = 0 To Height - 1
      d = Rnd()
      If d < (1 - p) Then
        mLandArr(i, j) = Mountain
      Else
        mLandArr(i, j) = Valley
      End If
    Next
  Next
  RaiseEvent Draw
End Sub
Public Sub Reset()
Dim i As Long, j As Long
  For i = 0 To Width - 1
    For j = 0 To Height - 1
      If mLandArr(i, j) And Valley Then
        mLandArr(i, j) = Valley
      End If
    Next
  Next
  RaiseEvent Draw
End Sub
Public Sub ClearFindTrack()
Dim i As Long, j As Long
  For i = 0 To Width - 1
    For j = 0 To Height - 1
      If mLandArr(i, j) = TrackFind Then
        mLandArr(i, j) = Valley
      End If
    Next
  Next
  RaiseEvent Draw
End Sub
Public Sub SaveFile(aFNr As Long)
  Dim i As Long, ni As Long: ni = UBound(mLandArr, 1)
  Dim j As Long, nj As Long: nj = UBound(mLandArr, 2)
  Print #aFNr, CStr(ni + 1)
  Print #aFNr, CStr(nj + 1)
  For j = 0 To nj
    For i = 0 To ni
      If i < ni Then Print #aFNr, GetStr(i, j); _
                Else Print #aFNr, GetStr(i, j)
    Next
  Next
End Sub
Public Sub LoadFile(aFNr As Long)
TryE: On Error GoTo CatchE
  Dim s As String
  Dim i As Long, ni As Long: Line Input #aFNr, s: ni = CLng(s) - 1
  Dim j As Long, nj As Long: Line Input #aFNr, s: nj = CLng(s) - 1
  ReDim mLandArr(0 To ni, 0 To nj)
  Do While Not EOF(aFNr) 'For i = 0 To ni
    Line Input #aFNr, s
    For i = 0 To Len(s) - 1
      If j <= nj And i <= ni Then _
        mLandArr(i, j) = GetLsE(Mid$(s, i + 1, 1))
    Next
    j = j + 1
  Loop
  Exit Sub
CatchE:
  Call ErrHandler("LoadFile")
End Sub
Public Sub ComputeCoords(aPB As PictureBox, ByVal SngX As Single, ByVal SngY As Single, ByRef LngX As Long, ByRef LngY As Long)
  Dim stepX As Double: stepX = aPB.ScaleWidth / (UBound(mLandArr, 1) + 1)
  Dim stepY As Double: stepY = aPB.ScaleHeight / (UBound(mLandArr, 2) + 1)
  LngX = Int(SngX / stepX)
  LngY = Int(SngY / stepY)
End Sub
Public Function CheckBounds(ByVal X As Long, Y As Long) As Boolean
  If 0 <= X And 0 <= Y Then
    If X <= UBound(mLandArr, 1) And Y <= UBound(mLandArr, 2) Then CheckBounds = True
  End If
End Function
Public Sub Invalidate()
  RaiseEvent Draw
End Sub
Public Sub Draw(aPB As PictureBox)
  Dim i As Long, j As Long
  Dim nx As Long: nx = UBound(mLandArr, 1)
  Dim ny As Long: ny = UBound(mLandArr, 2)
  Dim stepX As Double: stepX = aPB.ScaleWidth / (nx + 1)
  Dim stepY As Double: stepY = aPB.ScaleHeight / (ny + 1)
  Dim X As Double, Y As Double
  Dim colr As Long
  For j = 0 To ny
    For i = 0 To nx
      colr = GetColor(i, j)
      aPB.Line (X, Y)-(X + stepX, Y + stepY), colr, BF
      X = X + stepX
    Next
    Y = Y + stepY
    X = 0
  Next
End Sub
Public Sub DrawBlock(aPB As PictureBox, aX As Long, aY As Long)
TryE: On Error GoTo CatchE
  Dim stepX As Double: stepX = aPB.ScaleWidth / (UBound(mLandArr, 1) + 1)
  Dim stepY As Double: stepY = aPB.ScaleHeight / (UBound(mLandArr, 2) + 1)
  Dim X As Double: X = aX * stepX
  Dim Y As Double: Y = aY * stepY
  Dim colr As Long: colr = GetColor(aX, aY)
  aPB.Line (X, Y)-(X + stepX - 15, Y + stepY - 15), colr, BF
  Exit Sub
CatchE:
  Call ErrHandler("DrawBlock")
End Sub

'##############################'  Public Properties  '##############################'
Public Property Get Width() As Long
  Width = UBound(mLandArr, 1) + 1
End Property
Public Property Get Height() As Long
  Height = UBound(mLandArr, 2) + 1
End Property
Public Function IsVallay(X As Long, Y As Long) As Boolean
  IsVallay = (mLandArr(X, Y) = Valley)
End Function
Public Function IsMountain(X As Long, Y As Long) As Boolean
  IsMountain = (mLandArr(X, Y) = Mountain)
End Function
Public Property Get Point(X As Long, Y As Long) As LandscapeElement
  Point = mLandArr(X, Y)
End Property
Public Property Let Point(X As Long, Y As Long, RHS As LandscapeElement)
  mLandArr(X, Y) = RHS
  RaiseEvent DrawBlock(X, Y)
End Property

'##############################'  Private Procedures  '##############################'
Private Function GetStr(i As Long, j As Long) As String
  Select Case mLandArr(i, j)
  Case Mountain:  GetStr = "1"
  Case Valley:    GetStr = "0"
  Case InitPoint: GetStr = "2"
  Case GoalPoint: GetStr = "3"
  Case Else:      GetStr = "0"
  End Select
End Function
Private Function GetLsE(s As String) As LandscapeElement
  If s = "0" Then
    GetLsE = Valley
  ElseIf s = "1" Then
    GetLsE = Mountain
  ElseIf s = "2" Then
    GetLsE = InitPoint
  ElseIf s = "3" Then
    GetLsE = GoalPoint
  End If
End Function
Private Function GetColor(X As Long, Y As Long) As Long
  Select Case mLandArr(X, Y)
  Case Valley:    GetColor = &H9000&  'Dunkelgrün
  Case Mountain:  GetColor = &H909090 'Grau
  Case TrackBad:  GetColor = &HFF     'Rot
  Case TrackFind: GetColor = vbYellow '&H909090
  Case TrackGood: GetColor = vbGreen  '&H909090
  Case InitPoint: GetColor = vbBlue   'Blue
  Case GoalPoint: GetColor = vbWhite  'White'Black
  End Select
End Function
'Private Function CheckLandArr() As Boolean
'wird nimma gebraucht
'  CheckLandArr = ArrPtr(mLandArr) <> 0
'End Function

'##############################'   Private ErrHandler   '##############################'
Private Sub ErrHandler(aProcName As String, Optional Addinfo As String)
  Call GlobalErrHandler("Landscape", aProcName, Addinfo)
End Sub


