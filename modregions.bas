Attribute VB_Name = "modRegions"
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_MAX = RGN_COPY
Public Const RGN_MIN = RGN_AND
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_EX_TRANSPARENT = &H20&

Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Sub CutHole(This As Object, X As Long, Y As Long, Radius As Long)
Dim NewRGN As Long
Dim SquareRGN As Long
Dim CircleRGN As Long
Dim OtherRGN As Long
SquareRGN = CreateRectRgn(0, 0, This.Width, This.Height)
CircleRGN = CreateEllipticRgn(X - Radius, Y - Radius, X + Radius, Y + Radius)
OtherRGN = CreateEllipticRgn(170, 170, 200, 200)
'NewRGN = CircleRGN
NewRGN = SquareRGN
'NewRGN = OtherRGN
'CombineRgn NewRGN, SquareRGN, CircleRGN, RGN_DIFF
CombineRgn NewRGN, NewRGN, CircleRGN, RGN_DIFF
'CombineRgn NewRGN, CircleRGN, SquareRGN, RGN_DIFF
SetWindowRgn This.hWnd, NewRGN, True
DeleteObject NewRGN
DeleteObject SquareRGN
DeleteObject CircleRGN
End Sub

Function CutEdges(MyHWND As Long, Left, Top, Right, Bottom)
Dim MyRect As RECT
MyRect.Top = Top
MyRect.Bottom = Bottom
MyRect.Left = Left
MyRect.Right = Right
SetWindowRgn MyHWND, CreateEllipticRgnIndirect(MyRect), True
End Function

Sub CutRR(ThisObj As Object, Rad)
Set ThisObj = ThisObj
CutRoundRect ThisObj.hWnd, ThisObj.Width / 15, ThisObj.Height / 15, 0, 0, (ThisObj.Width / 15) / Rad, (ThisObj.Width / 15) / Rad
End Sub

Function CutCirCle(MyHWND As Long, Left, Top, Fat, Tall)
SetWindowRgn MyHWND, CreateEllipticRgn(Left, Top, Fat, Tall), True
End Function



Function CutRoundRect(MyHWND As Long, x1, y1, X2, Y2, X3, Y3)
SetWindowRgn MyHWND, CreateRoundRectRgn(x1, y1, X2, Y2, X3, Y3), True
End Function

Function CutRect(MyHWND As Long, x1, y1, X2, Y2)
SetWindowRgn MyHWND, CreateRectRgn(x1, y1, X2, Y2), True
End Function
Function CutPoly(MyHWND As Long, ByRef MyPoint As POINTAPI, Num)
SetWindowRgn MyHWND, CreatePolygonRgn(MyPoint, Num, 0), True
End Function





   



