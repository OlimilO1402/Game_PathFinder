VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PathFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '2007_06_14 Zeilen: 258
Public Enum EOrientation
  Compass = 0 '0000
  North = 1   '0001
  West = 2    '0010
  East = 4    '0100
  South = 8   '1000
  'NorthWest = 3 '1 + 2
  'NorthEast = 5 '1 + 5
  'SouthWest = 10 '8 + 2
  'SouthEast = 12 '8 + 4
End Enum
'Startposition
Private mInitPt      As LngPoint2D
'Endposition
Private mGoalPt      As LngPoint2D
Private mLand        As Landscape
Private mUseCompass  As Boolean
Private mClockwise   As Boolean 'im oder gegen den Uhrzeigersinn
Private mAnimDelay   As Long
Private mStepCounter As Long 'Schrittzähler
Private mCancel      As Boolean
Private mBeQuiet     As Boolean
Private mIsRunning   As Boolean
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub Class_Initialize()
  mAnimDelay = 50
  Call Reset
  mGoalPt.X = -1: mGoalPt.Y = -1
End Sub
Private Sub Class_Terminate()
  mCancel = True
  Set mLand = Nothing
End Sub

Public Sub NewC(aLand As Landscape)
  Set mLand = aLand
End Sub

Public Sub Reset()
  mStepCounter = 0
End Sub
Public Sub SetInitPos(X As Long, Y As Long)
  If mLand.CheckBounds(X, Y) Then
  'den alten InitPoint in mLand ausschalten
    If mLand.Point(mInitPt.X, mInitPt.Y) = InitPoint Then
      mLand.Point(mInitPt.X, mInitPt.Y) = Valley
    End If
    mInitPt.X = X:  mInitPt.Y = Y
    mLand.Point(mInitPt.X, mInitPt.Y) = mLand.Point(mInitPt.X, mInitPt.Y) Or InitPoint
  End If
End Sub
Public Sub GetInitPos(ByRef X As Long, ByRef Y As Long)
  X = mInitPt.X: Y = mInitPt.Y
End Sub
Public Sub SetGoalPos(X As Long, Y As Long)
  If mLand.CheckBounds(mGoalPt.X, mGoalPt.Y) Then
    'den alten GoalPoint in mLand ausschalten
    If mLand.Point(mGoalPt.X, mGoalPt.Y) = GoalPoint Then
      mLand.Point(mGoalPt.X, mGoalPt.Y) = Valley
    End If
  End If
  If mLand.CheckBounds(X, Y) Then
    mGoalPt.X = X:  mGoalPt.Y = Y
    mLand.Point(mGoalPt.X, mGoalPt.Y) = mLand.Point(mGoalPt.X, mGoalPt.Y) Or GoalPoint
  End If
End Sub
Public Sub GetGoalPos(ByRef X As Long, ByRef Y As Long)
  X = mGoalPt.X: Y = mGoalPt.Y
End Sub
Public Sub GoToGoal()
  Call Run(True)
End Sub
Public Sub WhichGoal()
  mGoalPt.X = -1: mGoalPt.Y = -1
  Call Run(False)
End Sub
Public Function StepCounterToString() As String
  StepCounterToString = CStr(mStepCounter)
End Function

Public Sub LoadFile(aFileName As String)
  Call FileStreamProc(aFileName, False)
End Sub
Public Sub SaveFile(aFileName As String)
  Call FileStreamProc(aFileName, True)
End Sub
Private Sub FileStreamProc(aFNm As String, ByVal bOut As Boolean)
TryE: On Error GoTo FinallyE
  Dim FNr As Long: FNr = FreeFile
  Dim l As String
  If bOut Then 'schreiben
    Open aFNm For Output As FNr
    Print #FNr, CStr(mInitPt.X):   Print #FNr, CStr(mInitPt.Y)
    Print #FNr, CStr(mGoalPt.X):   Print #FNr, CStr(mGoalPt.Y)
    Print #FNr, CStr(mUseCompass): Print #FNr, CStr(mClockwise)
    Print #FNr, CStr(mAnimDelay):  Call mLand.SaveFile(FNr)
  Else         'lesen
    Open aFNm For Input As FNr
    Line Input #FNr, l: mInitPt.X = CLng(l):    Line Input #FNr, l: mInitPt.Y = CLng(l)
    Line Input #FNr, l: mGoalPt.X = CLng(l):    Line Input #FNr, l: mGoalPt.Y = CLng(l)
    Line Input #FNr, l: mUseCompass = CBool(l): Line Input #FNr, l: mClockwise = CBool(l)
    Line Input #FNr, l: mAnimDelay = CLng(l):   Call mLand.LoadFile(FNr)
    mLand.Point(mInitPt.X, mInitPt.Y) = InitPoint
    mLand.Point(mGoalPt.X, mGoalPt.Y) = GoalPoint
  End If
FinallyE:
  Close FNr
  If Err = 0 Then Exit Sub
CatchE:
  Dim pn As String: pn = IIf(bOut, "SaveFile", "LoadFile") & "->FileStream"
  Call ErrHandler(pn)
End Sub
'##############################'  Public Properties  '##############################'
Public Property Get UseCompass() As Boolean:       UseCompass = mUseCompass: End Property
Public Property Let UseCompass(RHS As Boolean):   mUseCompass = RHS:         End Property

Public Property Get Clockwise() As Boolean:         Clockwise = mClockwise:  End Property
Public Property Let Clockwise(RHS As Boolean):     mClockwise = RHS:         End Property

Public Property Get AnimationDelay() As Long:  AnimationDelay = mAnimDelay:  End Property
Public Property Let AnimationDelay(RHS As Long):   mAnimDelay = RHS:         End Property

'und gleich wieder ausschalten
Public Property Get BeQuiet() As Boolean
  BeQuiet = mBeQuiet: mBeQuiet = False
End Property
Public Property Let BeQuiet(RHS As Boolean):  mBeQuiet = RHS: End Property

'und gleich wieder ausschalten
Public Property Get Cancel() As Boolean
  Cancel = mCancel:  mCancel = False
End Property
Public Property Let Cancel(RHS As Boolean)
  'If mIsRunning Then
  mCancel = RHS
End Property

Public Property Get IsRunning():  IsRunning = mIsRunning: End Property

'##############################'  Private Procedures  '##############################'
Private Sub Run(ByVal WithGoal As Boolean)
TryE: On Error GoTo FinallyE
  Dim mess As String
  Dim X As Long: X = mInitPt.X
  Dim Y As Long: Y = mInitPt.Y
  mStepCounter = 0
  mIsRunning = True
  If RecursiveFindPath(X, Y, WithGoal) Then
    If mCancel Then
      mess = "User hat Abbrechen gewählt!"
    Else
      mess = "OK Pfad gefunden! "
    End If
  Else
    If WithGoal Then
      mess = "The Goal was not reached! "
    Else
      mess = "All points that can be reached are red!"
    End If
    If Cancel Then mess = mess & vbCrLf & "User selected cancel"
  End If
  mess = mess & vbCrLf & " Steps: " & StepCounterToString
  If Not BeQuiet Then MsgBox mess
  Call mLand.Invalidate
FinallyE:
  mIsRunning = False
  If Err = 0 Then Exit Sub
CatchE:
  Call ErrHandler("Run")
End Sub
Private Function RecursiveFindPath(X As Long, Y As Long, WithGoal As Boolean) As Boolean
Dim bIsGoodPath As Boolean
Dim o As EOrientation
  If mCancel Then RecursiveFindPath = True: Exit Function
  If mAnimDelay >= 0 Then
    Call Sleep(mAnimDelay)
    DoEvents
  End If
  If Not mLand.CheckBounds(X, Y) Then
    RecursiveFindPath = False
    Exit Function
  End If
  mStepCounter = mStepCounter + 1
  If (X = mGoalPt.X) And (Y = mGoalPt.Y) Then
    RecursiveFindPath = True
    mLand.Point(X, Y) = GoalPoint
    If WithGoal Then Exit Function 'fertig
  End If
  Select Case mLand.Point(X, Y)
  Case Mountain, TrackFind, TrackBad
    RecursiveFindPath = False 'Wand oder Stelle schon erreicht
  Case Valley, InitPoint 'Stelle neu erreicht
    mLand.Point(X, Y) = TrackFind
    bIsGoodPath = False
    If mUseCompass Then
      o = GetCompassOrientation(X, Y)
    Else
      o = South
    End If
    Dim r As Byte 'Long
    Do While (Not bIsGoodPath) And r < 5  '(o < South + 1)
      Select Case o
      Case South: bIsGoodPath = RecursiveFindPath(X, Y + 1, WithGoal)
      Case East: bIsGoodPath = RecursiveFindPath(X + 1, Y, WithGoal)
      Case North: bIsGoodPath = RecursiveFindPath(X, Y - 1, WithGoal)
      Case West: bIsGoodPath = RecursiveFindPath(X - 1, Y, WithGoal)
      End Select
      If bIsGoodPath Then
        Exit Do 'Von Hier aus gibt es einen Weg ans Ziel.
      Else
        o = NextOrientation(o)
        r = r + 1
      End If
    Loop
    If Not bIsGoodPath Then
      mLand.Point(X, Y) = TrackBad
    End If
    RecursiveFindPath = bIsGoodPath
  End Select
End Function
Private Function GetCompassOrientation(ByVal X As Long, ByVal Y As Long) As EOrientation
  Dim o As EOrientation
  Dim dx As Long:   dx = mGoalPt.X - X 'Nord-Süd
  Dim dy As Long:   dy = mGoalPt.Y - Y 'West-Ost
  Dim dxy As Long: dxy = Abs(dx) - Abs(dy)
  If dxy > 0 Then 'West-Ost
    If dx >= 0 Then o = East Else o = West
  Else 'Nord-Süd
    If dy >= 0 Then o = South Else o = North
  End If
  GetCompassOrientation = o
End Function
Private Function NextOrientation(ByVal aO As EOrientation) As EOrientation
'hier könnte man eigentlich auch mal den Zufall entscheiden lassen,
'oder die nächste Richtung öfters mal den Kompass bestimmen lassen
  If mClockwise Then
    'im Uhrzeigersinn
    Select Case aO
    Case Compass
    Case North: NextOrientation = East
    Case East: NextOrientation = South
    Case South: NextOrientation = West
    Case West: NextOrientation = North
    End Select
  Else
    'gegen den Uhrzeigersinn
    Select Case aO
    Case Compass
    Case North: NextOrientation = West
    Case West: NextOrientation = South
    Case South: NextOrientation = East
    Case East: NextOrientation = North
    End Select
  End If
End Function

'##############################'   Private ErrHandler   '##############################'
Private Sub ErrHandler(aProcName As String, Optional Addinfo As String)
  Call GlobalErrHandler("PathFinder", aProcName, Addinfo)
End Sub

