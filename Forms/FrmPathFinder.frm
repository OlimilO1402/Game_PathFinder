VERSION 5.00
Begin VB.Form FrmPathFinder 
   Caption         =   "PathFinder"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   Icon            =   "FrmPathFinder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PbLandScape 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   3315
      TabIndex        =   16
      Top             =   840
      Width           =   3375
   End
   Begin VB.CheckBox ChkUseCompass 
      Caption         =   "Check2"
      Height          =   255
      Left            =   8160
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin VB.CheckBox ChkClockwise 
      Caption         =   "Check1"
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnLoadLand 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton BtnSaveLand 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton BtnPFGo 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton BtnWhichGoal 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton BtnGoalPoint 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnResetLand 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnRedoPoints 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnInitPoint 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton BtnValleyPoint 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnSetAnim 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton BtnMountainPoint 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BtnIterate 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton BtnNewLand 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmPathFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '2007_06_14 Zeilen: 219
Private Enum ClickPoints
  ClickNone = 0
  ClickInit = 1
  ClickGoal = 2
  ClickVall = 3
  ClickMntn = 4
End Enum
Private mClickPoint As ClickPoints
Private mInitPoint  As LngPoint2D
Private mGoalPoint  As LngPoint2D
Private mFileName   As String
Private WithEvents mLand As Landscape
Attribute mLand.VB_VarHelpID = -1
Private mPF As PathFinder

Private Sub ChkClockwise_Click()
  mPF.Clockwise = (ChkClockwise.Value = vbChecked)
End Sub

Private Sub Form_Load()
  Randomize
  Set mLand = New_Landscape(100, 75)
  Set mPF = New_PathFinder(mLand)
  ChkUseCompass.Value = mPF.UseCompass
  mPF.AnimationDelay = 1
  ChkUseCompass.Value = vbChecked
  BtnNewLand.Caption = "New Land"
  BtnResetLand.Caption = "Reset Land"
  BtnRedoPoints.Caption = "RedoPoints"
  BtnInitPoint.Caption = "Initial Point"
  BtnGoalPoint.Caption = "Goal Point"
  BtnValleyPoint.Caption = "Valley Point"
  BtnMountainPoint.Caption = "Mountain Pt"
  ChkClockwise.Caption = "Clockwise"
  BtnLoadLand.Caption = "Load Land"
  BtnSaveLand.Caption = "Save Land"
  BtnPFGo.Caption = "Go"
  BtnWhichGoal.Caption = "WhichGoal"
  BtnCancel.Caption = "Cancel"
  BtnSetAnim.Caption = "Anim. Speed"
  BtnIterate.Caption = "Clear Found"
  ChkUseCompass.Caption = "Compass"
End Sub
Private Sub Form_Unload(Cancel As Integer)
  mPF.BeQuiet = True
  mPF.Cancel = True
  Set mPF = Nothing
  Set mLand = Nothing
End Sub
Private Sub Form_Resize()
Dim l As Single, T As Single, W As Single, H As Single
Dim brdr As Single: brdr = 8 * 15
  l = brdr: T = brdr: W = 9 * brdr: H = 3 * brdr
  If W > 0 And H > 0 Then Call BtnNewLand.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnResetLand.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnRedoPoints.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnInitPoint.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnGoalPoint.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnValleyPoint.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnMountainPoint.Move(l, T, W, H)
  l = l + W + brdr
  If W > 0 And H > 0 Then Call ChkClockwise.Move(l, T, W, H)
  
  l = brdr: T = T + H
  If W > 0 And H > 0 Then Call BtnLoadLand.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnSaveLand.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnPFGo.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnWhichGoal.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnCancel.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnSetAnim.Move(l, T, W, H)
  l = l + W
  If W > 0 And H > 0 Then Call BtnIterate.Move(l, T, W, H)

  l = l + W + brdr
  If W > 0 And H > 0 Then Call ChkUseCompass.Move(l, T, W, H)
  
  l = brdr: T = T + H 'PbLandScape.Top
  W = Me.ScaleWidth - l - brdr
  H = Me.ScaleHeight - T - brdr
  If W > 0 And H > 0 Then Call PbLandScape.Move(l, T, W, H)
End Sub

Private Sub BtnNewLand_Click()
  Dim sw As String: sw = InputBox("Geben Sie die Breite an", "Breite?", CStr(mLand.Width))
  If (Len(sw) > 0) Then
    If IsNumeric(sw) Then
      Dim sh As String: sh = InputBox("Geben Sie die Höhe an", "Höhe?", CStr(mLand.Height))
      If Len(sh) > 0 Then
        Set mLand = New_Landscape(CLng(sw), CLng(sh))
        Set mPF = New_PathFinder(mLand)
        ChkUseCompass.Value = mPF.UseCompass
        Call PbLandScape.Refresh
        If IsNumeric(sh) Then
          Dim sp As String: sp = InputBox("Geben Sie einen Prozentwert für Täler/Gesamtfläche an.", "Täler-%", "0.67")
          If Len(sp) And IsNumeric(sp) Then
            Call mLand.GenerateRndLand(CDbl(sp))
            Exit Sub
          End If
        End If
      End If
    Else
      MsgBox "Please give a numeric value"
    End If
  End If
End Sub
Private Sub BtnResetLand_Click()
  mLand.Reset
  mPF.Cancel = False
End Sub
Private Sub BtnLoadLand_Click()
  If Len(mFileName) = 0 Then mFileName = App.Path & "\" & "Lab.pff"
  Dim FNm As String: FNm = InputBox("Geben Sie einen Dateinamen an:", "Datei öffnen", mFileName)
  If Len(FNm) Then
    mFileName = FNm
    Call mPF.LoadFile(FNm)
    Call mPF.GetInitPos(mInitPoint.X, mInitPoint.Y)
    Call mPF.GetGoalPos(mGoalPoint.X, mGoalPoint.Y)
    ChkUseCompass.Value = IIf(mPF.UseCompass, vbChecked, vbUnchecked)
    ChkClockwise.Value = IIf(mPF.Clockwise, vbChecked, vbUnchecked)
    Call mLand.Invalidate
  End If
End Sub
Private Sub BtnSaveLand_Click()
  If Len(mFileName) = 0 Then mFileName = App.Path & "\" & "Lab.pff"
  Dim FNm As String: FNm = InputBox("Geben Sie einen Dateinamen an:", "Datei speichern unter", mFileName)
  If Len(FNm) Then
    mFileName = FNm
    Call mPF.SaveFile(FNm)
  End If
End Sub
Private Sub BtnRedoPoints_Click()
  Call mPF.SetInitPos(mInitPoint.X, mInitPoint.Y)
  Call mPF.SetGoalPos(mGoalPoint.X, mGoalPoint.Y)
End Sub
Private Sub BtnInitPoint_Click()
  mClickPoint = ClickInit
End Sub
Private Sub BtnGoalPoint_Click()
  mClickPoint = ClickGoal
End Sub
Private Sub BtnValleyPoint_Click()
  mClickPoint = ClickVall
End Sub
Private Sub BtnMountainPoint_Click()
  mClickPoint = ClickMntn
End Sub

Private Sub BtnPFGo_Click()
  mPF.GoToGoal
End Sub
Private Sub BtnWhichGoal_Click()
  mPF.WhichGoal
End Sub
Private Sub BtnCancel_Click()
  mPF.Cancel = True
End Sub
Private Sub BtnSetAnim_Click()
Dim sa As String
  sa = InputBox("Animation delay [ms]", "Animation delay?", CStr(mPF.AnimationDelay))
  If Len(sa) And IsNumeric(sa) Then
    mPF.AnimationDelay = CLng(sa)
  End If
End Sub
Private Sub ChkUseCompass_Click()
  If Not mPF Is Nothing Then mPF.UseCompass = (ChkUseCompass.Value = vbChecked)
End Sub

Private Sub PbLandScape_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If mClickPoint <> ClickNone Then
    Dim LngX As Long, LngY As Long
    Call mLand.ComputeCoords(PbLandScape, X, Y, LngX, LngY)
    If mLand.IsVallay(LngX, LngY) Then
      Select Case mClickPoint
      Case ClickInit
        mInitPoint.X = LngX: mInitPoint.Y = LngY
        Call mPF.SetInitPos(mInitPoint.X, mInitPoint.Y)
        mClickPoint = ClickNone
      Case ClickGoal
        mGoalPoint.X = LngX: mGoalPoint.Y = LngY
        Call mPF.SetGoalPos(mGoalPoint.X, mGoalPoint.Y)
        mClickPoint = ClickNone
      Case ClickMntn
        mLand.Point(LngX, LngY) = Mountain
      End Select
    Else
      If mClickPoint = ClickVall Then
        mLand.Point(LngX, LngY) = Valley
      Else
        Dim mess As String
        mess = "Please select only green blocks" & vbCrLf & _
               "green = [valley, forrest, meadow]" & vbCrLf & _
               "grey = [mountains, rocks]"
        MsgBox mess
      End If
    End If
    'mClickPoint = ClickNone
  End If
End Sub

Private Sub mLand_Draw()
  Call mLand.Draw(PbLandScape)
End Sub
Private Sub mLand_DrawBlock(X As Long, Y As Long)
  Call mLand.DrawBlock(PbLandScape, X, Y)
End Sub
Private Sub PbLandScape_Paint()
  Call mLand.Draw(PbLandScape)
End Sub


