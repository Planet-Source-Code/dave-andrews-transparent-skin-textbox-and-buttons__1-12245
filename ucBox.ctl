VERSION 5.00
Begin VB.UserControl ucBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ScaleHeight     =   251
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   253
   ToolboxBitmap   =   "ucBox.ctx":0000
   Begin VB.PictureBox picPlace 
      AutoRedraw      =   -1  'True
      Height          =   135
      Left            =   1680
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   1320
      Width           =   135
   End
   Begin VB.HScrollBar hscrNotes 
      Height          =   135
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
   End
   Begin VB.VScrollBar vscrNotes 
      Height          =   1695
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   135
   End
   Begin VB.TextBox txtNotes 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2640
      Width           =   2175
   End
End
Attribute VB_Name = "ucBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim StartX As Single
Dim StartY As Single
Dim FinishX As Single
Dim FinishY As Single
Dim Painting As Boolean
Dim Scrolling As Boolean
Dim TTop As Integer
Dim TLeft As Integer
Dim Spos As Integer
Dim Epos As Integer
'Default Property Values:
Const m_def_CursorColor = 0
'Property Variables:
Dim m_CursorColor As Variant
Public Event TextChange(Text As String)



Private Sub HiLiteCursor()
Dim sX As Integer
Dim sY As Integer
Dim fX As Integer
Dim fY As Integer
Dim CsX As Integer
Dim CsY As Integer
Dim CfX As Integer
Dim CfY As Integer
Dim Lx As Integer
Dim Ly As Integer
Dim i As Integer
Dim OverLine As Boolean
'sX = -1
For i = 1 To txtNotes.SelStart
    If Mid(txtNotes.Text, i, 2) = vbNewLine Then
        sY = sY + 1
        sX = -1
    Else
        sX = sX + 1
    End If
Next i
If sX < 1 Then
    CsX = 0
    sX = 0
Else
    CsX = UserControl.TextWidth(Left(GetLine(sY), sX))
End If
CsY = sY * UserControl.TextHeight("X")
UserControl.DrawMode = vbInvert
For i = txtNotes.SelStart To txtNotes.SelStart + txtNotes.SelLength
    If i > 0 Then
        If Mid(txtNotes.Text, i, 2) = vbNewLine Then
            CfX = UserControl.TextWidth(Mid(GetLine(sY), sX + 1, fX + 1))
            CfY = UserControl.TextHeight("X")
            UserControl.Line (CsX, CsY)-(CsX + CfX, CsY + CfY), CursorColor, BF
            Lx = CsX + CfX
            Ly = CsY
            fY = sY
            sY = sY + 1
            CsX = 0
            CsY = CsY + CfY
            fX = -1
            sX = 0
            OverLine = True
        Else
            fX = fX + 1
        End If
    End If
Next i
If fX > 0 Then
    If OverLine Then
        CfX = UserControl.TextWidth(Mid(GetLine(sY), sX + 1, fX))
        CfY = UserControl.TextHeight("X")
    Else
        CfX = UserControl.TextWidth(Mid(GetLine(sY), sX + 1, fX - 1))
        CfY = UserControl.TextHeight("X")
    End If
    Lx = CsX + CfX
    Ly = CsY
    UserControl.Line (CsX, CsY)-(CsX + CfX, CsY + CfY), CursorColor, BF
End If
UserControl.DrawMode = vbCopyPen
UserControl.DrawWidth = 3
UserControl.Line (Lx, Ly)-(Lx, Ly + UserControl.TextHeight("X")), CursorColor
UserControl.DrawWidth = 1
If Ly + (2 * UserControl.TextHeight("X")) > TTop + UserControl.ScaleHeight Then TTop = Ly + (2 * UserControl.TextHeight("X")) - UserControl.ScaleHeight 'TTop + UserControl.TextHeight("X")
If Ly - (1 * UserControl.TextHeight("X")) <= TTop Then TTop = Ly - (1 * UserControl.TextHeight("X")) 'TTop - (2 * UserControl.TextHeight("X"))

If Lx + (2 * UserControl.TextWidth("X")) > TLeft + UserControl.ScaleWidth Then TLeft = Lx + (2 * UserControl.TextWidth("X")) - UserControl.ScaleWidth   'TLeft + UserControl.TextWidth("X")
If Lx - (2 * UserControl.TextWidth("X")) < TLeft Then TLeft = Lx - (2 * UserControl.TextWidth("X")) 'TLeft - (2 * UserControl.TextWidth("X"))

If TLeft < 0 Then TLeft = 0
If TTop < 0 Then TTop = 0
End Sub

Private Function GetLine(Line As Integer) As String
On Local Error Resume Next
Dim i As Integer
Dim cCount As Integer
Dim cStart As Integer
Dim lCount As Integer
cStart = 1
For i = 1 To Len(txtNotes.Text)
    If Mid(txtNotes.Text, i, 2) = vbNewLine Then
        'cCount = cCount + 2
        If Line = lCount Then
            Exit For
        Else
            cStart = cCount + 1
        End If
        lCount = lCount + 1
        
        'i = i + 2
    End If
    cCount = cCount + 1
Next i

GetLine = Mid(txtNotes.Text, cStart, cCount - cStart + 1)
If Left(GetLine, 2) = vbNewLine Then GetLine = Right(GetLine, Len(GetLine) - 2)
If Right(GetLine, 2) = vbNewLine Then GetLine = Left(GetLine, Len(GetLine) - 2)
End Function

Private Function lCount() As Integer
Dim i As Integer
For i = 1 To Len(txtNotes.Text)
    If Mid(txtNotes.Text, i, 2) = vbNewLine Then lCount = lCount + 1
Next i
End Function

Private Function MaxLine() As String
Dim i As Integer
Dim Max As String
Dim Start As Integer
Start = 1
For i = 1 To Len(txtNotes.Text)
    If Len(Max) < i - Start Then Max = Mid(txtNotes, Start, i - Start)
    If Mid(txtNotes, i, 2) = vbNewLine Then
        Start = i + 2
    End If
Next i
MaxLine = Max
End Function


Public Sub PaintText()
On Local Error Resume Next
Painting = True
If Not Scrolling Then
    Dim lSize As Integer
    Dim lMax As String
    lMax = MaxLine()
    lSize = lCount()
    vscrNotes.Max = lSize * UserControl.TextHeight("X") - UserControl.ScaleHeight + hscrNotes.Height + 2
    vscrNotes.Value = TTop
    hscrNotes.Max = UserControl.TextWidth(lMax) - UserControl.ScaleWidth + vscrNotes.Width + 2
    hscrNotes.Value = TLeft
    HiLiteCursor
End If
UserControl.Cls
UserControl.ScaleTop = TTop
UserControl.ScaleLeft = TLeft
UserControl.CurrentX = 0
UserControl.CurrentY = 0
UserControl.Print txtNotes.Text
HiLiteCursor
Painting = False
End Sub



Private Function GetSelPos(ByVal X As Integer, ByVal Y As Integer) As Integer
On Local Error GoTo SelAll
Dim i As Integer
Dim lCount As Integer
Dim cCount As Integer
Dim Line As Integer
Dim Char As Integer
Line = Fix(Y / UserControl.TextHeight("X"))
cCount = -1
For i = 1 To Len(txtNotes.Text)
    If Mid(txtNotes.Text, i, 2) = vbNewLine Then
        lCount = lCount + 1
        'cCount = cCount + 2
        'i = i + 2
    End If
    cCount = cCount + 1
    If lCount = Line Then
        If Line > 0 Then cCount = cCount + 2
        Exit For
    End If
Next i
If UserControl.TextWidth(GetLine(Line)) >= X Then
    i = 1
    X = X + (UserControl.TextWidth("X") / 2)
    Do While X > UserControl.TextWidth(Mid(txtNotes.Text, cCount + 1, i))
        If i > Len(txtNotes.Text) Then Exit Do
        i = i + 1
    Loop
Else
    i = Len(GetLine(Line)) + 1
End If
cCount = cCount + i - 1
SkipSel:
If cCount < 0 Then cCount = 0
GetSelPos = cCount
Exit Function
SelAll:
    cCount = Len(txtNotes.Text)
    GoTo SkipSel
End Function


Private Sub hscrNotes_Change()
If Painting Then Exit Sub
Scrolling = True
TLeft = hscrNotes
PaintText
Scrolling = False
End Sub

Private Sub hscrNotes_Scroll()
If Painting Then Exit Sub
Scrolling = True
TLeft = hscrNotes
PaintText
Scrolling = False
End Sub


Private Sub txtNotes_Change()
PaintText
RaiseEvent TextChange(txtNotes.Text)
End Sub

Private Sub txtNotes_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyLeft, vbKeyRight, vbKeyDown, vbKeyEnd, vbKeyHome
        PaintText
End Select
End Sub
Private Sub txtNotes_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyUp, vbKeyLeft, vbKeyRight, vbKeyDown, vbKeyEnd, vbKeyHome
        PaintText
End Select
End Sub


Private Sub UserControl_GotFocus()
txtNotes.SetFocus

End Sub

Private Sub UserControl_Initialize()
'PaintText txtNotes, picNotes
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And 1 Then
    txtNotes.SelStart = GetSelPos(X, Y)
    txtNotes.SelLength = 0
    txtNotes.SetFocus
    StartX = X
    StartY = Y
    PaintText
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And 1 Then
    Dim swap As Integer
    Spos = GetSelPos(StartX, StartY)
    Epos = GetSelPos(X, Y)
    If Spos > Epos Then
        swap = Epos
        Epos = Spos
        Spos = swap
        txtNotes.SelStart = Spos
    End If
    txtNotes.SelLength = Epos - Spos
    PaintText
End If
End Sub

Private Sub UserControl_Resize()
'If UserControl.BorderStyle = 1 Then
'    vscrNotes.Top = 1
'    vscrNotes.Left = UserControl.ScaleWidth - vscrNotes.Width - 1
'    vscrNotes.Height = UserControl.ScaleHeight - 2 - vscrNotes.Width
'    hscrNotes.Top = UserControl.ScaleHeight - hscrNotes.Height - 1
'    hscrNotes.Left = 1
'    hscrNotes.Width = UserControl.ScaleWidth - 2 - hscrNotes.Height
'Else
    vscrNotes.Top = 0
    vscrNotes.Left = UserControl.ScaleWidth - vscrNotes.Width
    vscrNotes.Height = UserControl.ScaleHeight - vscrNotes.Width
    hscrNotes.Top = UserControl.ScaleHeight - hscrNotes.Height
    hscrNotes.Left = 0
    hscrNotes.Width = UserControl.ScaleWidth - hscrNotes.Height
'End If
picPlace.Top = hscrNotes.Top
picPlace.Left = vscrNotes.Left
txtNotes.Width = 32000
txtNotes.Height = UserControl.Height - hscrNotes.Height
txtNotes.Left = UserControl.ScaleWidth + 100
PaintText
End Sub

Private Sub vscrNotes_Change()
If Painting Then Exit Sub
Scrolling = True
TTop = vscrNotes
PaintText
Scrolling = False

End Sub


Private Sub vscrNotes_Scroll()
If Painting Then Exit Sub
Scrolling = True
TTop = vscrNotes
PaintText
Scrolling = False

End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = UserControl.Image
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    vscrNotes.LargeChange = UserControl.TextHeight("X")
    vscrNotes.SmallChange = UserControl.TextHeight("X")
    hscrNotes.LargeChange = UserControl.TextWidth("X")
    hscrNotes.SmallChange = UserControl.TextWidth("X")
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get CursorColor() As Variant
    CursorColor = m_CursorColor
End Property

Public Property Let CursorColor(ByVal New_CursorColor As Variant)
    m_CursorColor = New_CursorColor
    PropertyChanged "CursorColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
    m_CursorColor = m_def_CursorColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_CursorColor = PropBag.ReadProperty("CursorColor", m_def_CursorColor)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 251)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 3)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 253)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    txtNotes.Text = PropBag.ReadProperty("Text", "")
    txtNotes.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtNotes.SelText = PropBag.ReadProperty("SelText", "")
    txtNotes.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtNotes.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtNotes.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtNotes.SelText = PropBag.ReadProperty("SelText", "")
    txtNotes.Text = PropBag.ReadProperty("Text", "")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("CursorColor", m_CursorColor, m_def_CursorColor)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 251)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 3)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 253)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("Text", txtNotes.Text, "")
    Call PropBag.WriteProperty("SelLength", txtNotes.SelLength, 0)
    Call PropBag.WriteProperty("SelText", txtNotes.SelText, "")
    Call PropBag.WriteProperty("SelStart", txtNotes.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtNotes.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtNotes.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtNotes.SelText, "")
    Call PropBag.WriteProperty("Text", txtNotes.Text, "")
End Sub
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ScaleHeight
'Public Property Get ScaleHeight() As Single
'    ScaleHeight = UserControl.ScaleHeight
'End Property
'
'Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
'    UserControl.ScaleHeight() = New_ScaleHeight
'    PropertyChanged "ScaleHeight"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ScaleLeft
'Public Property Get ScaleLeft() As Single
'    ScaleLeft = UserControl.ScaleLeft
'End Property
'
'Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
'    UserControl.ScaleLeft() = New_ScaleLeft
'    PropertyChanged "ScaleLeft"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ScaleMode
'Public Property Get ScaleMode() As Integer
'    ScaleMode = UserControl.ScaleMode
'End Property
'
'Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
'    UserControl.ScaleMode() = New_ScaleMode
'    PropertyChanged "ScaleMode"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ScaleTop
'Public Property Get ScaleTop() As Single
'    ScaleTop = UserControl.ScaleTop
'End Property
'
'Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
'    UserControl.ScaleTop() = New_ScaleTop
'    PropertyChanged "ScaleTop"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,ScaleWidth
'Public Property Get ScaleWidth() As Single
'    ScaleWidth = UserControl.ScaleWidth
'End Property
'
'Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
'    UserControl.ScaleWidth() = New_ScaleWidth
'    PropertyChanged "ScaleWidth"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtNotes,txtNotes,-1,Text
'Public Property Get Text() As String
'    Text = txtNotes.Text
'End Property
'
'Public Property Let Text(ByVal New_Text As String)
'    txtNotes.Text() = New_Text
'    PropertyChanged "Text"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtNotes,txtNotes,-1,SelLength
'Public Property Get SelLength() As Long
'    SelLength = txtNotes.SelLength
'End Property
'
'Public Property Let SelLength(ByVal New_SelLength As Long)
'    txtNotes.SelLength() = New_SelLength
'    PropertyChanged "SelLength"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtNotes,txtNotes,-1,SelText
'Public Property Get SelText() As String
'    SelText = txtNotes.SelText
'End Property
'
'Public Property Let SelText(ByVal New_SelText As String)
'    txtNotes.SelText() = New_SelText
'    PropertyChanged "SelText"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=txtNotes,txtNotes,-1,SelStart
'Public Property Get SelStart() As Long
'    SelStart = txtNotes.SelStart
'End Property
'
'Public Property Let SelStart(ByVal New_SelStart As Long)
'    txtNotes.SelStart() = New_SelStart
'    PropertyChanged "SelStart"
'End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNotes,txtNotes,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = txtNotes.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtNotes.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNotes,txtNotes,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtNotes.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtNotes.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNotes,txtNotes,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = txtNotes.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtNotes.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtNotes,txtNotes,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtNotes.Text
    PaintText
End Property

Public Property Let Text(ByVal New_Text As String)
    txtNotes.Text() = New_Text
    PropertyChanged "Text"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=dataNotes,dataNotes,-1,DatabaseName
'Public Property Get DatabaseName() As String
'    DatabaseName = dataNotes.DatabaseName
'End Property
'
'Public Property Let DatabaseName(ByVal New_DatabaseName As String)
'    dataNotes.DatabaseName() = New_DatabaseName
'    PropertyChanged "DatabaseName"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=dataNotes,dataNotes,-1,Database
'Public Property Get Database() As Database
'    Set Database = dataNotes.Database
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=dataNotes,dataNotes,-1,RecordSource
'Public Property Get RecordSource() As String
'    RecordSource = dataNotes.RecordSource
'End Property
'
'Public Property Let RecordSource(ByVal New_RecordSource As String)
'    dataNotes.RecordSource() = New_RecordSource
'    PropertyChanged "RecordSource"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=dataNotes,dataNotes,-1,Recordset
'Public Property Get Recordset() As Recordset
'    Set Recordset = dataNotes.Recordset
'End Property
'
'Public Property Set Recordset(ByVal New_Recordset As Recordset)
'    Set dataNotes.Recordset = New_Recordset
'    PropertyChanged "Recordset"
'End Property
'
