VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'This code researched and developed by Dave Andrews
'Check out my website:- http://www.audiokingdom.com  (independent music network)
'I know you probably don't care, but I thought I'de put a plug in for my company :)

'Feel free to use this wherever you want,
'I would just appreciate and credit / mention in your code
'-------------------------------------------------------------
Const COLOR_CAPTIONTEXT = 9
Const DT_CENTER = &H1      'centre left to right
Const DT_VCENTER = &H4     'centre top to bottom
Const DT_NOCLIP = &H100    'fast draw
Const DT_SINGLELINE = &H20 'single line only
Const DT_WORDBREAK = &H10

'Const DT_FLAGS = DT_SINGLELINE Or DT_CENTER Or DT_VCENTER Or DT_NOCLIP
Const DT_FLAGS = DT_CENTER Or DT_NOCLIP Or DT_VCENTER Or DT_WORDBREAK

Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long

'-----------------------------------------------------------
Dim Rounded As Boolean
Dim Clear As Boolean
Dim Level As Single
Dim Edge As Integer
Dim MyText As String
Dim Rad As Integer
Dim picHDC As Long
Dim conHDC As Long
Dim picHWND As Long
Dim bW As Integer
Dim bH As Integer
Dim bL As Integer
Dim bT As Integer
Dim FColor As Long
Dim BColor As Long
Dim dX As Integer
Dim dY As Integer
'--------HSL CONVERSION STUFF FOR BLENDING PIXELS-----------
'HSL conversion routines taken from Dan Redding's "Color Lab"
'Originally converted from the Microsoft Knowledge Base

Private Const HSLMAX As Integer = 240 '***
Const RGBMAX As Integer = 255 '***
Const UNDEFINED As Integer = (HSLMAX * 2 / 3) '***
Private Type HSLCol
    Hue As Integer
    Sat As Integer
    Lum As Integer
End Type
'-----------------------API DECLARATIONS-----------------------------------

Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal color As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long


Private Const SRCCOPY = &HCC0020


Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type






Private Sub EmbossText(tX As Integer, tY As Integer)
    Dim wTextParams As DRAWTEXTPARAMS
    Dim RC As RECT
    With RC
        .left = dX + Edge + 2 + tX
        .top = dY + Edge + tY
        .right = bW - Edge + 1
        .bottom = bH - Edge
    End With
    wTextParams.cbSize = Len(wTextParams)
    Call SetTextColor(picHDC, Brighten(GetPixel(picHDC, bW / 2, bH / 2), 0.4))
    Call DrawTextEx(picHDC, MyText, Len(MyText), RC, DT_FLAGS, wTextParams)
    With RC
        .left = dX + Edge + 1 + tX
        .top = dY + Edge - 1 + tY
        .right = bW - Edge
        .bottom = bH - Edge - 1
    End With
    wTextParams.cbSize = Len(wTextParams)
    Call SetTextColor(picHDC, FColor)
    Call DrawTextEx(picHDC, MyText, Len(MyText), RC, DT_FLAGS, wTextParams)

End Sub








Private Sub RaiseBevel()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim sPix As Long
Dim fPix As Long
Dim M As Single
'If transparent, then copy container image as background,
If Clear Then
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL + 1, bT + 1, vbSrcCopy
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL + 1, bT + 1, vbSrcCopy
Else 'we copy the container content onto the holder image and the backcolor
    RefreshHDC
End If 'Left Edge
M = 1
For i = 0 To Edge
    For j = i To bH - i
        sPix = Brighten(GetPixel(picHDC, i, j), M * Level)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel picHDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / Edge)
Next i
For j = Edge To bH - Edge
    SetPixel picHDC, Edge, j, Brighten(GetPixel(picHDC, Edge, j), Level)
Next j
'Top Edge
M = 1
For i = 0 To Edge
    For j = i To bW - i
        sPix = Brighten(GetPixel(picHDC, j, i), M * Level)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel picHDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / Edge)
Next i
For j = Edge To bW - Edge
    SetPixel picHDC, j, i, Brighten(GetPixel(picHDC, j, Edge), Level)
Next j
'Right Edge
M = 0
For i = bW - Edge To bW
    For j = bW - i To bH - (bW - i)
        sPix = Darken(GetPixel(picHDC, i, j), M * Level)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel picHDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / Edge)
Next i
For j = Edge To bH - Edge
    SetPixel picHDC, (bW - Edge), j, Darken(GetPixel(picHDC, (bW - Edge), j), Level)
Next j
'Bottom Edge
M = 0
For i = bH - Edge To bH
    For j = bH - i To bW - (bH - i)
        sPix = Darken(GetPixel(picHDC, j, i), M * Level)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel picHDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / Edge)
Next i
For j = Edge To bW - Edge
    SetPixel picHDC, j, (bH - Edge), Darken(GetPixel(picHDC, j, (bH - Edge)), Level)
Next j

End Sub

Private Sub InsetBevel()
'Creates a bevel around the perimeter of an object
'By blending the outside pixels of the objects container
'on a gradient scale.
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim sPix As Long
Dim fPix As Long
Dim M As Single
'If transparent, then copy container image as background,
If Clear Then
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL - 1, bT - 1, vbSrcCopy
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL - 1, bT - 1, vbSrcCopy
Else 'we copy the container content onto the holder image and the backcolor
    RefreshHDC
End If 'Left Edge
M = 1
For i = 0 To Edge - 1
    For j = i To bH - i
        sPix = Darken(GetPixel(picHDC, i, j), M * Level)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel picHDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / Edge)
Next i
For j = Edge To bH - Edge
    SetPixel picHDC, Edge, j, Darken(GetPixel(picHDC, Edge, j), Level)
Next j
'Top Edge
M = 1
For i = 0 To Edge - 1
    For j = i To bW - i
        sPix = Darken(GetPixel(picHDC, j, i), M * Level)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel picHDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / Edge)
Next i
For j = Edge To bW - Edge
    SetPixel picHDC, j, i, Darken(GetPixel(picHDC, j, Edge), Level)
Next j
'Right Edge
M = 0
For i = (bW - Edge) + 1 To bW
    For j = bW - i To bH - (bW - i)
        sPix = Brighten(GetPixel(picHDC, i, j), M * Level)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel picHDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / Edge)
Next i
For j = Edge To bH - Edge
    SetPixel picHDC, (bW - Edge), j, Brighten(GetPixel(picHDC, (bW - Edge), j), Level)
Next j
'Bottom Edge
M = 0
For i = (bH - Edge) + 1 To bH
    For j = bH - i To bW - (bH - i)
        sPix = Brighten(GetPixel(picHDC, j, i), M * Level)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel picHDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / Edge)
Next i
For j = Edge To bW - Edge
    SetPixel picHDC, j, (bH - Edge), Brighten(GetPixel(picHDC, j, (bH - Edge)), Level)
Next j

End Sub

Private Sub RaiseRound()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim sPix As Long
Dim fPix As Long
Dim M As Single
Dim s As Single
Dim pi
pi = 4 * Atn(1)
'If transparent, then copy container image as background,
If Clear Then
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL + 1, bT + 1, vbSrcCopy
Else 'we copy the container content onto the holder image and the backcolor
    RefreshHDC
End If
'Left Edge
'GoTo Corners
M = 1
For i = 0 To Edge - 1
    For j = (Rad + Edge) To bH - (Rad + Edge)
        sPix = Brighten(GetPixel(picHDC, i, j), M * Level)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel picHDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / Edge)
Next i
For j = (Rad + Edge) To bH - (Rad + Edge)
    sPix = Brighten(GetPixel(picHDC, Edge, j), Level)
    SetPixel picHDC, Edge, j, sPix
Next j
'Top Edge
M = 1
For i = 0 To Edge - 1
    For j = Rad + Edge To bW - (Rad + Edge)
        sPix = Brighten(GetPixel(picHDC, j, i), M * Level)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel picHDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / Edge)
Next i
For j = Rad + Edge To bW - (Rad + Edge)
    sPix = Brighten(GetPixel(picHDC, j, Edge), Level)
    SetPixel picHDC, j, Edge, sPix
Next j
'Right Edge
M = 0
For i = (bW - Edge) + 1 To bW
    For j = Rad + Edge To bH - (Rad + Edge)
        sPix = Darken(GetPixel(picHDC, i, j), M * Level)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel picHDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / Edge)
Next i
For j = Rad + Edge To bH - (Rad + Edge)
    sPix = Darken(GetPixel(picHDC, (bW - Edge), j), Level)
    SetPixel picHDC, (bW - Edge), j, sPix
Next j
'Bottom Edge
M = 0
For i = (bH - Edge) + 1 To bH
    For j = Rad + Edge To bW - (Rad + Edge)
        sPix = Darken(GetPixel(picHDC, j, i), M * Level)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel picHDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / Edge)
Next i
For j = Rad + Edge To bW - (Rad + Edge)
    sPix = Darken(GetPixel(picHDC, j, (bH - Edge)), Level)
    SetPixel picHDC, j, (bH - Edge), sPix
Next j
'--------------CORNERS-----------------------
Corners:
Dim rX As Integer
Dim rY As Integer
Dim lX As Integer
Dim lY As Integer
Dim rStep As Integer
'Top Left Corder
M = 0
For i = Rad To Rad + Edge
    For j = 91 To 179
        rX = (i * Cos(j * (pi / 180))) + (Rad + Edge)
        rY = -(i * Sin(j * (pi / 180))) + (Rad + Edge)
        If lX <> rX Or lY <> rY Then
            If i = Rad Then
                sPix = Brighten(GetPixel(picHDC, rX, rY), Level)
            Else
                sPix = Brighten(GetPixel(picHDC, rX, rY), M * Level)
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel picHDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
    Next j
    M = M + (1 / (Edge))
Next i
'Top Right Corder
M = 0
For i = Rad To Rad + Edge
    s = 0
    For j = 91 To 179
        rX = -(i * Cos(j * (pi / 180))) + (bW - Rad - Edge)
        rY = -(i * Sin(j * (pi / 180))) + (Rad + Edge)
        If lX <> rX Or lY <> rY Then
            If j < 135 Then
                If i = Rad Then
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (1 - s) * Level)
                Else
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (M - s) * Level)
                End If
            Else
                If i = Rad Then
                    sPix = Darken(GetPixel(picHDC, rX, rY), (1 - s) * Level)
                Else
                    sPix = Darken(GetPixel(picHDC, rX, rY), (M - s) * Level)
                End If
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel picHDC, rX, rY, Blend(sPix, fPix, M)
        End If
        If j < 135 Then
            s = s + (1 / 45)
        Else
            s = s - (1 / 45)
        End If
        lX = rX
        lY = rY

    Next j
    M = M + (1 / (Edge))
Next i
'Bottom Left Corder
M = 0
For i = Rad To Rad + Edge
    s = 0
    For j = 91 To 179
        rX = (i * Cos(j * (pi / 180))) + (Rad + Edge)
        rY = (i * Sin(j * (pi / 180))) + (bH - Rad - Edge)
        If lX <> rX Or lY <> rY Then
            If j > 135 Then
                If i = Rad Then
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (1 + s) * Level)
                Else
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (M + s) * Level)
                End If
            Else
                If i = Rad Then
                    sPix = Darken(GetPixel(picHDC, rX, rY), (1 + s) * Level)
                Else
                    sPix = Darken(GetPixel(picHDC, rX, rY), (M + s) * Level)
                End If
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel picHDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
        If j > 135 Then
            s = s + (1 / 45)
        Else
            s = s - (1 / 45)
        End If
    Next j
    M = M + (1 / (Edge))
Next i
'Bottom right Corder
M = 0
For i = Rad To Rad + Edge
    For j = 91 To 179
        rX = -(i * Cos(j * (pi / 180))) + (bW - Rad - Edge)
        rY = (i * Sin(j * (pi / 180))) + (bH - Rad - Edge)
        If lX <> rX Or lY <> rY Then
            If i = Rad Then
                sPix = Darken(GetPixel(picHDC, rX, rY), Level)
            Else
                sPix = Darken(GetPixel(picHDC, rX, rY), M * Level)
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel picHDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
    Next j
    M = M + (1 / (Edge))
Next i

End Sub

Private Sub InsetRound()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim sPix As Long
Dim fPix As Long
Dim M As Single
Dim s As Single
Dim pi
pi = 4 * Atn(1)
'If transparent, then copy container image as background,
If Clear Then
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL - 1, bT - 1, vbSrcCopy
    BitBlt picHDC, 0, 0, bW, bH, conHDC, bL - 1, bT - 1, vbSrcCopy
Else 'we copy the container content onto the holder image and the backcolor
    RefreshHDC
End If 'Left Edge
'GoTo Corners
M = 1
For i = 0 To Edge - 1
    For j = (Rad + Edge) To bH - (Rad + Edge)
        sPix = Darken(GetPixel(picHDC, i, j), M * Level)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel picHDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / Edge)
Next i
For j = (Rad + Edge) To bH - (Rad + Edge)
    sPix = Darken(GetPixel(picHDC, Edge, j), Level)
    SetPixel picHDC, Edge, j, sPix
Next j
'Top Edge
M = 1
For i = 0 To Edge - 1
    For j = Rad + Edge To bW - (Rad + Edge)
        sPix = Darken(GetPixel(picHDC, j, i), M * Level)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel picHDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M - (1 / Edge)
Next i
For j = Rad + Edge To bW - (Rad + Edge)
    sPix = Darken(GetPixel(picHDC, j, Edge), Level)
    SetPixel picHDC, j, Edge, sPix
Next j
'Right Edge
M = 0
For i = (bW - Edge) + 1 To bW
    For j = Rad + Edge To bH - (Rad + Edge)
        sPix = Brighten(GetPixel(picHDC, i, j), M * Level)
        fPix = GetPixel(conHDC, i + bL, j + bT)
        SetPixel picHDC, i, j, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / Edge)
Next i
For j = Rad + Edge To bH - (Rad + Edge)
    sPix = Brighten(GetPixel(picHDC, (bW - Edge), j), Level)
    SetPixel picHDC, (bW - Edge), j, sPix
Next j
'Bottom Edge
M = 0
For i = (bH - Edge) + 1 To bH
    For j = Rad + Edge To bW - (Rad + Edge)
        sPix = Brighten(GetPixel(picHDC, j, i), M * Level)
        fPix = GetPixel(conHDC, j + bL, i + bT)
        SetPixel picHDC, j, i, Blend(sPix, fPix, M)
    Next j
    M = M + (1 / Edge)
Next i
For j = Rad + Edge To bW - (Rad + Edge)
    sPix = Brighten(GetPixel(picHDC, j, (bH - Edge)), Level)
    SetPixel picHDC, j, (bH - Edge), sPix
Next j
'--------------CORNERS-----------------------
Corners:
'rad = rad + 1 'return the value to it's original settings
Dim rX As Integer
Dim rY As Integer
Dim lX As Integer
Dim lY As Integer
Dim rStep As Integer
'Top Left Corder
M = 0
For i = Rad To Rad + Edge
    For j = 91 To 179
        rX = (i * Cos(j * (pi / 180))) + (Rad + Edge)
        rY = -(i * Sin(j * (pi / 180))) + (Rad + Edge)
        If lX <> rX Or lY <> rY Then
            If i = Rad Then
                sPix = Darken(GetPixel(picHDC, rX, rY), Level)
            Else
                sPix = Darken(GetPixel(picHDC, rX, rY), M * Level)
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel picHDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
    Next j
    M = M + (1 / (Edge))
Next i
'Top Right Corder
M = 0
For i = Rad To Rad + Edge
    s = 0
    For j = 91 To 179
        rX = -(i * Cos(j * (pi / 180))) + (bW - Rad - Edge)
        rY = -(i * Sin(j * (pi / 180))) + (Rad + Edge)
        If lX <> rX Or lY <> rY Then
            If j < 135 Then
                If i = Rad Then
                    sPix = Darken(GetPixel(picHDC, rX, rY), (1 - s) * Level)
                Else
                    sPix = Darken(GetPixel(picHDC, rX, rY), (M - s) * Level)
                End If
            Else
                If i = Rad Then
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (1 - s) * Level)
                Else
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (M - s) * Level)
                End If
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel picHDC, rX, rY, Blend(sPix, fPix, M)
        End If
        If j < 135 Then
            s = s + (1 / 45)
        Else
            s = s - (1 / 45)
        End If
        lX = rX
        lY = rY

    Next j
    M = M + (1 / (Edge))
Next i
'Bottom Left Corder
M = 0
For i = Rad To Rad + Edge
    s = 0
    For j = 91 To 179
        rX = (i * Cos(j * (pi / 180))) + (Rad + Edge)
        rY = (i * Sin(j * (pi / 180))) + (bH - Rad - Edge)
        If lX <> rX Or lY <> rY Then
            If j > 135 Then
                If i = Rad Then
                    sPix = Darken(GetPixel(picHDC, rX, rY), (1 + s) * Level)
                Else
                    sPix = Darken(GetPixel(picHDC, rX, rY), (M + s) * Level)
                End If
            Else
                If i = Rad Then
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (1 + s) * Level)
                Else
                    sPix = Brighten(GetPixel(picHDC, rX, rY), (M + s) * Level)
                End If
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel picHDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
        If j > 135 Then
            s = s + (1 / 45)
        Else
            s = s - (1 / 45)
        End If
    Next j
    M = M + (1 / (Edge))
Next i
'Bottom right Corder
M = 0
For i = Rad To Rad + Edge
    For j = 91 To 179
        rX = -(i * Cos(j * (pi / 180))) + (bW - Rad - Edge)
        rY = (i * Sin(j * (pi / 180))) + (bH - Rad - Edge)
        If lX <> rX Or lY <> rY Then
            If i = Rad Then
                sPix = Brighten(GetPixel(picHDC, rX, rY), Level)
            Else
                sPix = Brighten(GetPixel(picHDC, rX, rY), M * Level)
            End If
            fPix = GetPixel(conHDC, rX + bL, rY + bT)
            SetPixel picHDC, rX, rY, Blend(sPix, fPix, M)
        End If
        lX = rX
        lY = rY
    Next j
    M = M + (1 / (Edge))
Next i

End Sub




Function InitButton(pTarget As PictureBox, lText As String, bRounded As Boolean, bLevel As Single, bEdge As Integer, bRad As Integer, bClear As Boolean, Optional tX As Integer, Optional tY As Integer) As Single
pTarget.Container.ScaleMode = vbPixels
pTarget.ScaleMode = vbPixels
pTarget.Container.AutoRedraw = True
pTarget.AutoRedraw = True
bW = pTarget.ScaleWidth - 1
bH = pTarget.ScaleHeight - 1
bL = pTarget.left
bT = pTarget.top
dX = tX
dY = tY
picHDC = pTarget.hDC
conHDC = pTarget.Container.hDC
picHWND = pTarget.hWnd
Level = bLevel
Edge = bEdge
MyText = lText
Rounded = bRounded
Clear = bClear
Rad = bRad
FColor = pTarget.ForeColor
BColor = pTarget.BackColor
'--------------CREATE A REGION IF IT'S ROUNDED----------------------------
If bRounded Then
    Dim NewRGN As Long
    If bW = bH Then
        NewRGN = CreateEllipticRgn(0, 0, bW + 1, bH + 1)
    Else
        NewRGN = CreateEllipticRgn(0, 0, ((bRad + bEdge) * 2) + 1, ((bRad + bEdge) * 2) + 1)
        CombineRgn NewRGN, NewRGN, CreateEllipticRgn(bW + 1, bH + 1, (bW - 1) - ((bRad + bEdge) * 2), (bH - 1) - ((bRad + bEdge) * 2)), 2
        CombineRgn NewRGN, NewRGN, CreateEllipticRgn(0, bH + 1, ((bRad + bEdge) * 2) + 1, (bH - 1) - ((bRad + bEdge) * 2)), 2
        CombineRgn NewRGN, NewRGN, CreateEllipticRgn(bW + 1, 0, (bW - 1) - ((bRad + bEdge) * 2), ((bRad + bEdge) * 2) + 1), 2
        CombineRgn NewRGN, NewRGN, CreateRectRgn(0, (bRad + bEdge) + 1, bW + 1, bH - (bRad + bEdge)), 2
        CombineRgn NewRGN, NewRGN, CreateRectRgn((bRad + bEdge) + 1, 0, bW - (bRad + bEdge), bH + 1), 2
    End If
    SetWindowRgn pTarget.hWnd, NewRGN, True
    '-----------Get Images for rounded button
    RaiseRound
Else
    RaiseBevel
End If
EmbossText 0, 0
RefreshHDC
End Function




Private Function NZ(ZInput, Optional ZDefault) As Variant
On Error GoTo SkipIt:
If IsMissing(ZDefault) Then ZDefault = ""
If (IsNull(ZInput)) Or (ZInput = Empty) Or (ZInput = "") Then
    NZ = ZDefault
Else
    NZ = ZInput
End If
Exit Function
SkipIt:
NZ = ""
End Function



Private Function Brighten(RGBColor As Long, Percent As Single)
'Brightens a color by a decimal percent
Dim HSL As HSLCol, L As Long
    If Percent <= 0 Then
        Brighten = RGBColor
        Exit Function
    End If
    
    HSL = RGBtoHSL(RGBColor)
    L = HSL.Lum + (HSLMAX * Percent)
    If L > HSLMAX Then L = HSLMAX
    HSL.Lum = L
    Brighten = HSLtoRGB(HSL)
End Function
Private Function Darken(RGBColor As Long, Percent As Single)
'Darkens a color by a percent
Dim HSL As HSLCol, L As Long
    If Percent <= 0 Then
        Darken = RGBColor
        Exit Function
    End If
    
    HSL = RGBtoHSL(RGBColor)
    L = HSL.Lum - (HSLMAX * Percent)
    If L < 0 Then L = 0
    HSL.Lum = L
    Darken = HSLtoRGB(HSL)
End Function
Private Function Blend(RGB1 As Long, RGB2 As Long, Percent As Single) As Long
'blends two colors together by a certain percent (decimal percent)
Dim R As Integer, R1 As Integer, R2 As Integer, G As Integer, G1 As Integer, G2 As Integer, B As Integer, B1 As Integer, B2 As Integer
    
    If Percent >= 1 Then
        Blend = RGB2
        Exit Function
    ElseIf Percent <= 0 Then
        Blend = RGB1
        Exit Function
    End If
    
    R1 = RGBRed(RGB1)
    R2 = RGBRed(RGB2)
    G1 = RGBGreen(RGB1)
    G2 = RGBGreen(RGB2)
    B1 = RGBBlue(RGB1)
    B2 = RGBBlue(RGB2)
    
    R = ((R2 * Percent) + (R1 * (1 - Percent)))
    G = ((G2 * Percent) + (G1 * (1 - Percent)))
    B = ((B2 * Percent) + (B1 * (1 - Percent)))
    
    Blend = RGB(R, G, B)
End Function
Private Function iMax(a As Integer, B As Integer) _
    As Integer
'Return the Larger of two values
    iMax = IIf(a > B, a, B)
End Function

Private Function iMin(a As Integer, B As Integer) _
    As Integer
'Return the smaller of two values
    iMin = IIf(a < B, a, B)
End Function

Private Function RGBRed(RGBCol As Long) As Integer
If RGBCol = -1 Then Exit Function
'Return the Red component from an RGB Color
    RGBRed = RGBCol And &HFF
End Function

Private Function RGBGreen(RGBCol As Long) As Integer
If RGBCol = -1 Then Exit Function
'Return the Green component from an RGB Color
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Private Function RGBBlue(RGBCol As Long) As Integer
If RGBCol = -1 Then Exit Function
'Return the Blue component from an RGB Color
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function
Private Function HSLtoRGB(HueLumSat As HSLCol) As Long '***
'Converts HSL to a color value
    Dim R As Double, G As Double, B As Double
    Dim H As Double, L As Double, s As Double
    Dim Magic1 As Double, Magic2 As Double
    H = HueLumSat.Hue
    L = HueLumSat.Lum
    s = HueLumSat.Sat
    If CInt(s) = 0 Then
        R = (L * RGBMAX) / HSLMAX
        G = R
        B = R
    Else
        If L <= HSLMAX / 2 Then
            Magic2 = (L * (HSLMAX + s) + 0.5) / HSLMAX
        Else
            Magic2 = L + s - ((L * s) + 0.5) / HSLMAX
        End If
        Magic1 = 2 * L - Magic2
        R = (HuetoRGB(Magic1, Magic2, H + (HSLMAX / 3)) * RGBMAX + 0.5) / HSLMAX
        G = (HuetoRGB(Magic1, Magic2, H) * RGBMAX + 0.5) / HSLMAX
        B = (HuetoRGB(Magic1, Magic2, H - (HSLMAX / 3)) * RGBMAX + 0.5) / HSLMAX
    End If
    HSLtoRGB = RGB(CInt(R), CInt(G), CInt(B))
End Function

Private Function HuetoRGB(mag1 As Double, mag2 As Double, ByVal Hue As Double) As Double     '***
'Utility function for HSLtoRGB
    If Hue < 0 Then
        Hue = Hue + HSLMAX
    ElseIf Hue > HSLMAX Then
        Hue = Hue - HSLMAX
    End If
    Select Case Hue
        Case Is < (HSLMAX / 6)
            HuetoRGB = (mag1 + (((mag2 - mag1) * Hue + _
                (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Is < (HSLMAX / 2)
            HuetoRGB = mag2
        Case Is < (HSLMAX * 2 / 3)
            HuetoRGB = (mag1 + (((mag2 - mag1) * _
                ((HSLMAX * 2 / 3) - Hue) + _
                (HSLMAX / 12)) / (HSLMAX / 6)))
        Case Else
            HuetoRGB = mag1
    End Select
End Function
Private Function RGBtoHSL(RGBCol As Long) As HSLCol '***
'Returns an HSLCol datatype containing Hue, Luminescence
'and Saturation; given an RGB Color value

Dim R As Integer, G As Integer, B As Integer
Dim cMax As Integer, cMin As Integer
Dim RDelta As Double, GDelta As Double, _
    BDelta As Double
Dim H As Double, s As Double, L As Double
Dim cMinus As Long, cPlus As Long
    
    R = RGBRed(RGBCol)
    G = RGBGreen(RGBCol)
    B = RGBBlue(RGBCol)
    
    cMax = iMax(iMax(R, G), B) 'Highest and lowest
    cMin = iMin(iMin(R, G), B) 'color values
    
    cMinus = cMax - cMin 'Used to simplify the
    cPlus = cMax + cMin  'calculations somewhat.
    
    'Calculate luminescence (lightness)
    L = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)
    
    If cMax = cMin Then 'achromatic (r=g=b, greyscale)
        s = 0 'Saturation 0 for greyscale
        H = UNDEFINED 'Hue undefined for greyscale
    Else
        'Calculate color saturation
        If L <= (HSLMAX / 2) Then
            s = ((cMinus * HSLMAX) + 0.5) / cPlus
        Else
            s = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
        End If
    
        'Calculate hue
        RDelta = (((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus
        GDelta = (((cMax - G) * (HSLMAX / 6)) + 0.5) / cMinus
        BDelta = (((cMax - B) * (HSLMAX / 6)) + 0.5) / cMinus
    
        Select Case cMax
            Case CLng(R)
                H = BDelta - GDelta
            Case CLng(G)
                H = (HSLMAX / 3) + RDelta - BDelta
            Case CLng(B)
                H = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select
        
        If H < 0 Then H = H + HSLMAX
    End If
    
    RGBtoHSL.Hue = CInt(H)
    RGBtoHSL.Lum = CInt(L)
    RGBtoHSL.Sat = CInt(s)
End Function

Sub TriggerButton()
If Rounded Then
    InsetRound
Else
    InsetBevel
End If
EmbossText 1, 1
RefreshHDC
Do While picHWND = GetCapture()
    DoEvents
Loop
If Rounded Then
    RaiseRound
Else
    RaiseBevel
End If
EmbossText 0, 0
RefreshHDC
End Sub

Sub TriggerDown()
If Rounded Then
    InsetRound
Else
    InsetBevel
End If
EmbossText 1, 1
RefreshHDC
End Sub
Sub TriggerUp()
If Rounded Then
    RaiseRound
Else
    RaiseBevel
End If
EmbossText 0, 0
RefreshHDC
End Sub

Public Sub RefreshHDC()
    'Dim hBrush As Long
    Dim tRect As RECT

    'hBrush = CreateSolidBrush(BColor)
    With tRect
        .left = -1
        .top = -1
        .right = bW + 2
        .bottom = bH + 2
    End With

    'Call FillRect(picHDC, tRect, hBrush)
    'Call DeleteObject(hBrush)
    'DoEvents
    Call InvalidateRect(picHWND, tRect, False)
End Sub












