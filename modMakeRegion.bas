Attribute VB_Name = "modMakeRegion"
Option Explicit

Private Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, _
    ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" _
    (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
    ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" _
    (ByVal hWnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Long) As Long

Public Function fMakeATranspArea(frm As Form, AreaType As String, pCordinate() As Long) As Boolean
Const RGN_DIFF = 4

Dim lOriginalForm As Long, ltheHole As Long, lNewForm As Long, _
    lFwidth As Single, lFHeight As Single, lborder_width As Single, _
    ltitle_height As Single
On Error GoTo Trap
    lFwidth = frm.ScaleX(frm.Width, vbTwips, vbPixels)
    lFHeight = frm.ScaleY(frm.Height, vbTwips, vbPixels)
    lOriginalForm = CreateRectRgn(0, 0, lFwidth, lFHeight)
    lborder_width = (lFHeight - frm.ScaleWidth) / 2
    ltitle_height = lFHeight - lborder_width - frm.ScaleHeight

    Select Case AreaType
        Case "Elliptic"
            ltheHole = CreateEllipticRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
        Case "RectAngle"
            ltheHole = CreateRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4))
        Case "RoundRect"
            ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(5), pCordinate(6))
        Case "Circle"
            ltheHole = CreateRoundRectRgn(pCordinate(1), pCordinate(2), pCordinate(3), pCordinate(4), pCordinate(3), pCordinate(4))
        Case Else
            MsgBox "Unknown Shape!"
            Exit Function
    End Select

    lNewForm = CreateRectRgn(0, 0, 0, 0)
    CombineRgn lNewForm, lOriginalForm, ltheHole, 1
    SetWindowRgn frm.hWnd, lNewForm, True
    frm.Refresh
    fMakeATranspArea = False
Exit Function
Trap:
End Function



