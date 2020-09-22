VERSION 5.00
Begin VB.Form frmSmile 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSmile.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSmile.frx":27A2
   ScaleHeight     =   750
   ScaleWidth      =   750
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmSmile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)
'Window Positioning API
Private Declare Function SetWindowPos Lib "user32.dll" _
    (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1

Private Sub Form_Load()
Dim lParam(1 To 6) As Long
    GetPic
    lParam(1) = 1
    lParam(2) = 1
    lParam(3) = 50
    lParam(4) = 50
    lParam(5) = 50
    lParam(6) = 50
    fMakeATranspArea Me, "Circle", lParam()
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE
    Skitter
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE
    For Each frm In Forms
        Unload frm
    Next frm
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ((Shift And vbShiftMask) > 0) Then
    Me.Visible = False
    DoEvents
    PlayResSound 103, True
    Unload Me
Else
    Skitter
    If Fix(Rnd() * 3) = 2 Then GetPic
End If
End Sub

Private Sub Skitter()
Dim iRnd As Integer, sLoop As Single, lX As Long, lY As Long, lNewX As Long, lNewY As Long, Child As frmSmile

Static blnMoving As Boolean
    If Not blnMoving Then
        blnMoving = True
        Randomize
        iRnd = Fix(Rnd() * 20)
        Select Case iRnd
            Case 3
                PlayResSound 101, False
            Case 5
                PlayResSound 102, False
            Case 9
                PlayResSound 103, False
            Case 13
                PlayResSound 104, False
            Case 15
                PlayResSound 105, False
            Case 18
                PlayResSound 106, False
        End Select
        lX = Me.Left
        lY = Me.Top
        lNewX = Rnd() * (Screen.Width - 750)
        lNewY = Rnd() * (Screen.Height - 750)
        For sLoop = 1 To 0.05 Step -0.05
            Me.Move lNewX + ((lX - lNewX) * sLoop), lNewY + ((lY - lNewY) * sLoop)
            DoEvents
            Sleep 10
        Next sLoop
        Me.Move lNewX, lNewY
        
        If Fix(Rnd() * 25) = 1 Then
            Set Child = New frmSmile
            Load Child
            Child.Move Me.Left, Me.Top
            Child.Show
            Set Child = Nothing
        End If
        blnMoving = False
    End If
End Sub


Private Sub GetPic()
    If Fix(Rnd() * 8) = 2 Then
        Me.Picture = LoadResPic("G", 101)
    Else
        Me.Picture = LoadResPic("G", 102)
    End If
End Sub
