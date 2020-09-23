VERSION 5.00
Begin VB.UserControl ListBoxFX 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   DrawWidth       =   53
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   124
   Begin VB.PictureBox PicIcons 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   255
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox ImgBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   135
      ScaleHeight     =   345
      ScaleWidth      =   540
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1065
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox PicBase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      Begin VB.VScrollBar vBar 
         Height          =   270
         Left            =   570
         Max             =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   270
      End
   End
End
Attribute VB_Name = "ListBoxFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Simple Listbox Exmaple
' By DreamVB

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByRef lpPoint As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Enum GRADIENT_DIR
    Horizontal = &H0
    Vertical = &H1
End Enum

'Sort consts
Enum SortOrder
    Ascending = 1
    Descending = 2
End Enum

Enum AppearanceConstants
    Flat = 0
    C3D = 1
End Enum

'Text Alignments
Enum TextAlign
    aLeft = 0
    aCenter = 1
    aRight = 2
End Enum

'Listbox Border Style Consts
Enum bStyle
    bNone = 0
    bSingle = 1
End Enum

'Listbox Selectin styles
Enum SelectType
    sSoildColor = 0
    sBitmap
End Enum

'Listbox Background Styles
Enum BackGroundStyle
    bSoildColor = 0
    bBitmap = 1
    bGradient = 2
End Enum

'Picture Background alignments
Enum BackGroundAlign
    TopLeft = 0
    TopRight = 1
    Center = 2
    Tile = 3
    BottomLeft = 4
    BottomRight = 5
End Enum

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type TRIVERTEX
    x As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Private Type Item
    Item As Variant
    Key As Variant
    IconIdx As Integer
    Checked As Boolean
End Type

'Text Style Consts
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4

Private m_ItemData() As Item
Private m_ListCount As Long

Private m_ItemHeight As Integer, m_ItemWidth As Integer
Private m_ImageHeight As Integer, m_ImageWidth As Integer

Private m_StartColor As OLE_COLOR
Private m_EndColor As OLE_COLOR
Private m_Direction As GRADIENT_DIR

Private m_LastItemOffset As Long
Private m_ListIndex As Long

Private m_Created As Boolean
Private m_BackGround_Style As BackGroundStyle
Private m_PicAlignment As BackGroundAlign
Private m_TextAlignment As TextAlign
Private m_GridLines As Boolean
Private m_sort As SortOrder
Private m_FullRowSelect As Boolean
Private m_ShowIcons As Boolean

Private m_SelectionClolor As OLE_COLOR, m_ShowSelection As Boolean
Private m_TextColor As OLE_COLOR, m_SelectTxtColor As OLE_COLOR, m_IconMaskColor As OLE_COLOR

'Event Declarations:
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event Change()

Private Sub setTriVertexColor(tTV As TRIVERTEX, oColor As Long)
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    
    lRed = (oColor And &HFF&) * &H100&
    lGreen = (oColor And &HFF00&)
    lBlue = (oColor And &HFF0000) \ &H100&
    
    setTriVertexColorComponent tTV.Red, lRed
    setTriVertexColorComponent tTV.Green, lGreen
    setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent(ByRef oColor As Integer, ByVal lComponent As Long)
    If (lComponent And &H8000&) = &H8000& Then
        oColor = (lComponent And &H7F00&)
        oColor = oColor Or &H8000
    Else
        oColor = lComponent
    End If
End Sub

Private Sub GDI_GradientFill(hdc As Long, mRect As RECT, mStartColor As OLE_COLOR, mEndColor As OLE_COLOR, gDir As GRADIENT_DIR)
Dim gRect As GRADIENT_RECT
Dim tTV(0 To 1) As TRIVERTEX
    'Function used to paint a Gradient effect on the listbox
    
    setTriVertexColor tTV(1), TranslateColor(mEndColor)
    tTV(0).x = mRect.Left
    tTV(0).Y = mRect.top
    
    setTriVertexColor tTV(0), TranslateColor(mStartColor)
    tTV(1).x = mRect.Right
    tTV(1).Y = mRect.Bottom
    
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    
    GradientFill hdc, tTV(0), 2, gRect, 1, gDir
    
End Sub

Private Function Sort(sOrder As SortOrder)
Dim x As Long, Y As Long
    'Simple Insert sort
    For x = 1 To m_ListCount - 1
        For Y = x + 1 To m_ListCount - 1
            If sOrder = Ascending Then 'Sort Ascending
                If LCase(m_ItemData(x).Item) > LCase(m_ItemData(Y).Item) Then
                    Swap m_ItemData(x).Item, m_ItemData(Y).Item
                    Swap m_ItemData(x).Key, m_ItemData(Y).Key
                End If
            Else 'Sort Descending
                If LCase(m_ItemData(x).Item) < LCase(m_ItemData(Y).Item) Then
                    Swap m_ItemData(x).Item, m_ItemData(Y).Item
                    Swap m_ItemData(x).Key, m_ItemData(Y).Key
                End If
            End If
        Next Y
    Next x
    x = 0
    Y = 0
End Function

Private Sub Swap(a, b)
Dim t
    'Swaps two values
    t = b
    b = a
    a = t
End Sub

Private Sub GDI_DrawGrid()
Dim x As Long
    For x = 0 To PicBase.ScaleHeight
        GDI_LineTo PicBase.hdc, 0, (x * m_ItemHeight), CLng(m_ItemWidth), (x * m_ItemHeight), 1, vbButtonFace
    Next x
    x = 0
End Sub

Private Sub TileBackPattem()
Dim Align(3) As Long
Dim ImgH As Long, ImgW As Long
Dim rc As RECT
Dim hBrush As Long

    'Picture alignments
    ImgH = (ImgBack.ScaleHeight \ Screen.TwipsPerPixelY)
    ImgW = (ImgBack.ScaleWidth \ Screen.TwipsPerPixelX)
    h = (((UserControl.ScaleHeight \ Screen.TwipsPerPixelY) \ 2) - ImgW \ 2)
    
    Select Case m_PicAlignment
        Case TopLeft
            BitBlt PicBase.hdc, 0, 0, ImgW, ImgH, ImgBack.hdc, 0, 0, vbSrcCopy
        Case TopRight
            BitBlt PicBase.hdc, (m_ItemWidth - ImgW), 0, ImgW, ImgH, ImgBack.hdc, 0, 0, vbSrcCopy
        Case Center
            BitBlt PicBase.hdc, (m_ItemWidth - ImgW) \ 2, h, ImgW, ImgH, ImgBack.hdc, 0, 0, vbSrcCopy
        Case BottomLeft
            BitBlt PicBase.hdc, 0, (PicBase.ScaleHeight - ImgH), ImgW, ImgH, ImgBack.hdc, 0, 0, vbSrcCopy
        Case BottomRight
            BitBlt PicBase.hdc, (m_ItemWidth - ImgW), (PicBase.ScaleHeight - ImgH), ImgW, ImgH, ImgBack.hdc, 0, 0, vbSrcCopy
        Case Tile
            rc.Left = 0: rc.Right = m_ItemWidth
            rc.top = 0: rc.Bottom = PicBase.ScaleHeight
            hBrush = CreatePatternBrush(ImgBack.Picture)
            FillRect PicBase.hdc, rc, hBrush
            DeleteObject hBrush
    End Select

End Sub

Private Function SelectItem(ItemY As Long, hdc As Long, oColor As OLE_COLOR)
Dim rc As RECT
Dim hBrush As Long, bExtra As Integer

    With rc
        .Left = 0
        If m_FullRowSelect Then
            .Right = m_ItemWidth
        Else
            If (m_ShowIcons) Then bExtra = m_ImageWidth Else bExtra = 0
            .Right = PicBase.TextWidth(m_ItemData(m_ListIndex).Item) + PicBase.TextWidth(" ") + bExtra + 1
        End If
        
        .top = (ItemY + m_ItemHeight) + 1
        .Bottom = 1 + ItemY
    End With
    
    hBrush = CreateSolidBrush(TranslateColor(oColor))
    FillRect hdc, rc, hBrush
    DeleteObject hBrush
    
End Function

Private Function PrintText(ItemY As Long, hdc As Long, lText As String)
Dim rc As RECT
    With rc
        .Left = PicBase.CurrentX
        .Right = m_ItemWidth
        .top = (ItemY + m_ItemHeight)
        .Bottom = ItemY
    End With
    
    DrawText hdc, lText + " ", Len(lText) + 1, rc, DT_SINGLELINE Or m_TextAlignment Or DT_VCENTER

End Function

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Private Sub GDI_LineTo(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional LnWidth As Long = 1, Optional LineColor As OLE_COLOR)
Dim hPen As Long

    'Used for line drawing same as the one in VB with some extras like drawwidth and draw style
    hPen = CreatePen(0, LnWidth, TranslateColor(LineColor)) 'Create a soild pen
    DeleteObject SelectObject(hdc, hPen)   ' select the DC to draw onto
    If X1 >= 0 Then MoveToEx hdc, X1, Y1, 0
    LineTo hdc, X2, Y2 'Draw the line
    
End Sub

Public Sub AddItem(Item As String, Optional Key As Variant = "", Optional Icon As Integer = -1)
    m_ListCount = m_ListCount + 1
    ReDim Preserve m_ItemData(m_ListCount)
    
    m_ItemData(m_ListCount).Item = Item
    m_ItemData(m_ListCount - 1).Key = Key
    m_ItemData(m_ListCount).IconIdx = Icon
    
    
    If Not m_Created Then
        m_Created = True
    End If
   
    Call RenderItem(m_ListCount)
    PicBase.Refresh
End Sub

Public Sub AppendItem(Index As Long, Item As String)
On Error GoTo FlagErr:
Dim Temp As String
    
    Temp = m_ItemData(Index).Item
    Temp = Temp & Item
    m_ItemData(Index).Item = Temp
    Temp = ""
    Call RenderLB
    Exit Sub
FlagErr:
    If Err Then Err.Raise 9 + vbObjectError
End Sub

Public Sub AppendKey(Index As Long, Key As Variant)
On Error GoTo FlagErr:
Dim Temp As String
    
    Temp = m_ItemData(Index).Key
    Temp = Temp & Key
    m_ItemData(Index).Key = Temp
    Temp = ""
    Call RenderLB
    Exit Sub
FlagErr:
    If Err Then Err.Raise 9 + vbObjectError
End Sub

Public Sub Clear()
    m_ListCount = 0
    m_Created = False
    Erase m_ItemData
    PicBase.Cls
    Call RenderLB
End Sub

Public Sub Delete(Index As Long)
On Error GoTo ErrFlag:
Dim iSize As Long
    iSize = UBound(m_ItemData)
    
    'Deletes an list item
    If (iSize = 0) Then
        Clear
        Exit Sub
    End If
    
    If (Index > ListCount) Then
        Err.Raise 9
        Exit Sub
    End If
    
    While (iSize > Index)
        m_ItemData(Index).Item = m_ItemData(Index + 1).Item
        m_ItemData(Index).Key = m_ItemData(Index + 1).Key
        m_ItemData(Index).IconIdx = m_ItemData(Index + 1).IconIdx
        Index = Index + 1
    Wend
    
    ReDim Preserve m_ItemData(iSize - 1)
    m_ListCount = m_ListCount - 1
    iSize = 0
    
    Call RenderLB
    Exit Sub
    
ErrFlag:
If Err Then Err.Raise 9 + vbObjectError
End Sub

Public Sub MoveEx(CurIndex As Long, NewIndex As Long)
On Error GoTo ErrFlag:
    'Swaps item,key,icon to a new position
    If (CurIndex < 0) Or (CurIndex > m_ListCount) Or _
    (NewIndex < 0) Or (NewIndex > m_ListCount) Then
        Err.Raise 9
        Exit Sub
    End If
    
    Swap m_ItemData(CurIndex).Item, m_ItemData(NewIndex).Item
    Swap m_ItemData(CurIndex).Key, m_ItemData(NewIndex).Key
    Swap m_ItemData(CurIndex).IconIdx, m_ItemData(NewIndex).IconIdx
    Call RenderLB
    
Exit Sub
ErrFlag:
If Err Then Err.Raise 9 + vbObjectError
End Sub

Public Function ItemEquals(Item As String) As Boolean
Dim x As Long
    'Returns True if an item is equal to m_ItemData.item
    For x = 0 To m_ListCount - 1
        If m_ItemData(x).Item = Item Then
            ItemEquals = True
            Exit For
        End If
    Next
    x = 0
End Function

Public Function KeyEquals(Key As Variant) As Boolean
Dim x As Long
    'Returns True if an key is equal to m_ItemData.key
    For x = 0 To m_ListCount - 1
        If m_ItemData(x).Key = Key Then
            KeyEquals = True
            Exit For
        End If
    Next
    x = 0
End Function

Private Sub RenderItem(Item As Long)
Dim ItemY_Pos As Long
Dim MoveX As Integer
Dim rc As RECT

    On Error Resume Next
    
    If (m_ListCount) > m_LastItemOffset Then
        vBar.Max = (m_ListCount - m_LastItemOffset)
    Else
        vBar.Max = 0
    End If
    
    ItemY_Pos = (Item - vBar.Value - 1) * m_ItemHeight
    
    With m_ItemData(Item)
        If (Item = m_ListIndex) Then
            PicBase.ForeColor = m_SelectTxtColor
            If m_ShowSelection Then
                SelectItem ItemY_Pos, PicBase.hdc, m_SelectionClolor
            End If
        Else
            PicBase.ForeColor = m_TextColor
        End If
        
        If (m_ShowIcons) Then
            MoveX = m_ImageWidth + 2
        Else
            MoveX = 2
        End If
        
        If (PicIcons.Picture) <> 0 And m_ShowIcons = True <> 0 Then
            TransparentBlt PicBase.hdc, 0, ItemY_Pos + (m_ImageHeight \ 2) - 7, m_ImageWidth, m_ImageHeight, PicIcons.hdc, _
            .IconIdx * m_ImageHeight, 0, m_ImageWidth, m_ImageHeight, m_IconMaskColor
        End If
        
        PicBase.CurrentX = MoveX
        Call PrintText(ItemY_Pos, PicBase.hdc, CStr(.Item))
    End With
    
End Sub

Private Sub RenderLB()
Dim sPos As Long, xPos As Long, rc As RECT
On Error Resume Next

    PicBase.Cls
    'Last item offset to draw to
    m_LastItemOffset = Fix(PicBase.ScaleHeight / m_ItemHeight) ' + 1
    
    sPos = (vBar.Value + 1)
    xPos = (sPos - 1)
    
    'Add the Textured Background
    If (m_BackGround_Style = bBitmap) Then
        If (ImgBack.Picture) <> 0 And UserControl.Ambient.UserMode Then
            TileBackPattem
        End If
    End If
    
    'Add Gradient background
    If (m_BackGround_Style = bGradient) And UserControl.Ambient.UserMode Then
       rc.Left = 0: rc.Right = m_ItemWidth
       rc.top = 0: rc.Bottom = PicBase.ScaleHeight
       Call GDI_GradientFill(PicBase.hdc, rc, m_StartColor, m_EndColor, m_Direction)
    End If
    
    'Render each of the list items
    Do While (xPos < m_ListCount)
        xPos = xPos + 1
        Call RenderItem(xPos)
        If (xPos > sPos + m_LastItemOffset + 1) Then Exit Do
    Loop
    
    'Draw the GridLines
    If (m_GridLines) And UserControl.Ambient.UserMode Then
        Call GDI_DrawGrid
    End If
    
    xPos = 0
    sPos = 0
End Sub

Private Sub ResizeAll()
    PicBase.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    vBar.Move (PicBase.ScaleWidth - vBar.Width), 0, vBar.Width, PicBase.ScaleHeight
    m_ItemWidth = (PicBase.ScaleWidth - vBar.Width)
End Sub

Private Sub PicBase_Click()
    RaiseEvent Click
    RaiseEvent Change
End Sub

Private Sub PicBase_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    
Start:
    Select Case KeyCode
        Case vbKeyHome
            ListIndex = 0
            vBar.Value = 0
        Case vbKeyEnd
            ListIndex = m_ListCount - 1
            vBar.Value = vBar.Max
        Case vbKeyRight
            KeyCode = vbKeyDown
            GoTo Start:
        Case vbKeyLeft
            KeyCode = vbKeyUp
            GoTo Start:
        Case vbKeyUp
            If (ListIndex = 0) Then Exit Sub
            ListIndex = ListIndex - 1
            If Not ItemInScope(ListIndex) And (vBar.Value > 0) Then _
            vBar.Value = vBar.Value - 1
            RaiseEvent Change
        Case vbKeyDown
            If (m_ListIndex + 1 >= m_ListCount + 1) Then Exit Sub
            ListIndex = ListIndex + 1
            If Not ItemInScope(ListIndex) And (vBar.Value < vBar.Max) Then _
            vBar.Value = vBar.Value + 1
            RaiseEvent Change
    End Select
    
End Sub

Private Function ItemInScope(Index As Long) As Boolean
Dim sPos As Long
    sPos = (PicBase.ScaleHeight \ m_ItemHeight)
    If (Index > vBar.Value) And (Index < vBar.Value + sPos) Then
        ItemInScope = True
    End If
End Function

Private Sub PicBase_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim Pre_Idx As Integer
    RaiseEvent MouseDown(Button, Shift, x, Y)
    
    Pre_Idx = m_ListIndex
    If Button <> vbLeftButton Then Exit Sub
    m_ListIndex = ((Y \ m_ItemHeight) + vBar.Value) + 1
    If (m_ListIndex <= 0) Or (m_ListIndex > m_ListCount) Then m_ListIndex = Pre_Idx
   ' MsgBox m_ItemWidth \ 15
    Call RenderLB
    
End Sub

Private Sub UserControl_InitProperties()
    m_ImageHeight = 16
    m_ImageWidth = 16
    m_ShowSelection = True
    m_FullRowSelect = True
    m_GridLines = False
    m_ShowIcons = False
    m_sort = Ascending
    m_IconMaskColor = RGB(255, 0, 255)
    m_SelectionClolor = vbHighlight
    m_SelectTxtColor = vbWhite
    m_EndColor = vbWhite
    m_GridColor = vbButtonFace
    m_Direction = Vertical
    m_TextColor = vbBlack
    m_PicAlignment = TopLeft
    m_TextAlignment = aLeft
    m_BackGround_Style = bSoildColor
End Sub

Private Sub UserControl_Paint()
    Call RenderLB
End Sub

Private Sub UserControl_Resize()
    Call ResizeAll
End Sub

Private Sub UserControl_Show()
Dim aText_Height As Integer
    aText_Height = PicBase.TextHeight("A")
    
    If Not m_ShowIcons Then
        m_ItemHeight = aText_Height
    Else
        If (aText_Height > m_ImageHeight) Then
            m_ImageHeight = aText_Height
        End If
        m_ItemHeight = m_ImageHeight
    End If
    
    Sort Sorted
End Sub

Private Sub vBar_Change()
    RenderLB
End Sub

Private Sub vBar_Scroll()
    vBar_Change
End Sub

'Control Properties stuff
Public Property Get ListCount() As Long
    ListCount = m_ListCount
End Property

Public Property Get ItemSelectedText() As String
On Error Resume Next
    'Returns text select Items text
    ItemSelectedText = m_ItemData(m_ListIndex).Item
End Property

Public Property Get ListItem(Index As Long) As String
    'Return ListItem Data
    ListItem = m_ItemData(Index).Item
End Property

Public Property Let ListItem(Index As Long, vItem As String)
    'Set ListItem Data
    m_ItemData(Index).Item = vItem
    Call RenderLB
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
    'Return ListItem Index
    ListIndex = (m_ListIndex - 1)
End Property

Public Property Let ListIndex(Index As Long)
On Error GoTo ErrFlag:
    
    If (Index < 0) Then
        Err.Raise 9
        Exit Property
    End If

    If (Index = 0) Then
        m_ListIndex = 1
    Else
        m_ListIndex = Index + 1
    End If
    
    Call RenderLB
    RaiseEvent Change
    Exit Property
ErrFlag:
    If Err Then Err.Raise 9 + vbObjectError
End Property

Public Property Get Key(Index As Long) As Variant
    'Return ListItem Key
    Key = m_ItemData(Index).Key
End Property

Public Property Let Key(Index As Long, vKey As Variant)
    'Set ListItem Key
    m_ItemData(Index).Key = vKey
    Call RenderLB
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = PicBase.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    PicBase.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get SelectionClolor() As OLE_COLOR
    SelectionClolor = m_SelectionClolor
End Property

Public Property Let SelectionClolor(ByVal vNewSelect As OLE_COLOR)
    m_SelectionClolor = vNewSelect
Call RenderLB
    PropertyChanged "SelectionClolor"
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

Public Property Let TextColor(ByVal vNewTxtColor As OLE_COLOR)
    m_TextColor = vNewTxtColor
    Call RenderLB
    PropertyChanged "TextColor"
End Property

Public Property Get TextSelectColor() As OLE_COLOR
   TextSelectColor = m_SelectTxtColor
End Property

Public Property Let TextSelectColor(ByVal vNewSelTxtColr As OLE_COLOR)
    m_SelectTxtColor = vNewSelTxtColr
    PropertyChanged "TextSelectColor"
    Call RenderLB
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    PicBase.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    PicBase.Enabled = PropBag.ReadProperty("Enabled", True)
    Set PicBase.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_SelectionClolor = PropBag.ReadProperty("SelectionClolor", vbHighlight)
    m_TextColor = PropBag.ReadProperty("TextColor", 0)
    m_SelectTxtColor = PropBag.ReadProperty("TextSelectColor", vbWhite)
    m_ShowSelection = PropBag.ReadProperty("ShowSelection", True)
    Set PicBase.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackGround_Style = PropBag.ReadProperty("BackGroundStyle", bSoildColor)
    m_GridLines = PropBag.ReadProperty("GridLines", False)
    m_PicAlignment = PropBag.ReadProperty("BackGroundAlignment", TopLeft)
    m_FullRowSelect = PropBag.ReadProperty("FullRowSelect", True)
    Set BackImage = PropBag.ReadProperty("BackImage", Nothing)
    PicBase.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    PicBase.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    m_sort = PropBag.ReadProperty("Sorted", Ascending)
    PicBase.ToolTipText = PropBag.ReadProperty("Hint", "")
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    m_IconMaskColor = PropBag.ReadProperty("IconMaskColor", &HFF00FF)
    Set PicIcons.Picture = PropBag.ReadProperty("IconStrip", Nothing)
    m_ShowIcons = PropBag.ReadProperty("ShowIcons", False)
    m_ImageHeight = PropBag.ReadProperty("ImageHeight", 16)
    m_ImageWidth = PropBag.ReadProperty("ImageWidth", 16)
    m_TextAlignment = PropBag.ReadProperty("TextAlignment", aLeft)
    m_StartColor = PropBag.ReadProperty("gStartColor", vbBlue)
    m_EndColor = PropBag.ReadProperty("gEndColor", vbWhite)
    m_Direction = PropBag.ReadProperty("gDirection", &H1)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", PicBase.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", PicBase.Enabled, True)
    Call PropBag.WriteProperty("Font", PicBase.Font, Ambient.Font)
    Call PropBag.WriteProperty("TextColor", m_TextColor, 0)
    Call PropBag.WriteProperty("TextSelectColor", m_SelectTxtColor, vbWhite)
    Call PropBag.WriteProperty("SelectionClolor", m_SelectionClolor, vbHighlight)
    Call PropBag.WriteProperty("Font", PicBase.Font, Ambient.Font)
    Call PropBag.WriteProperty("ShowSelection", m_ShowSelection, True)
    Call PropBag.WriteProperty("BackGroundStyle", m_BackGround_Style, bSoildColor)
    Call PropBag.WriteProperty("BackImage", BackImage, Nothing)
    Call PropBag.WriteProperty("BackGroundAlignment", m_PicAlignment, TopLeft)
    Call PropBag.WriteProperty("FullRowSelect", m_FullRowSelect, True)
    Call PropBag.WriteProperty("GridLines", m_GridLines, False)
    Call PropBag.WriteProperty("Enabled", PicBase.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", PicBase.MousePointer, 0)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("Sorted", m_sort, Ascending)
    Call PropBag.WriteProperty("Hint", PicBase.ToolTipText, "")
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("IconMaskColor", m_IconMaskColor, &HFF00FF)
    Call PropBag.WriteProperty("IconStrip", PicIcons.Picture, Nothing)
    Call PropBag.WriteProperty("ShowIcons", m_ShowIcons, False)
    Call PropBag.WriteProperty("ImageHeight", m_ImageHeight, 16)
    Call PropBag.WriteProperty("ImageWidth", m_ImageWidth, 16)
    Call PropBag.WriteProperty("TextAlignment", m_TextAlignment, aLeft)
    Call PropBag.WriteProperty("gStartColor", m_StartColor, vbBlue)
    Call PropBag.WriteProperty("gEndColor", m_EndColor, vbWhite)
    Call PropBag.WriteProperty("gDirection", m_Direction, &H1)

End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = PicBase.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set PicBase.Font = New_Font
    PropertyChanged "Font"
    Call RenderLB
End Property

Public Property Get ShowSelection() As Boolean
    ShowSelection = m_ShowSelection
End Property

Public Property Let ShowSelection(ByVal New_Select As Boolean)
    m_ShowSelection = New_Select
End Property

Public Property Get BackGroundStyle() As BackGroundStyle
   BackGroundStyle = m_BackGround_Style
End Property

Public Property Let BackGroundStyle(ByVal New_BackGround As BackGroundStyle)
    m_BackGround_Style = New_BackGround
    PropertyChanged "BackGroundStyle"
    Call RenderLB
End Property

Public Property Get BackImage() As Picture
Attribute BackImage.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set BackImage = ImgBack.Picture
End Property

Public Property Set BackImage(ByVal New_BackImage As Picture)
    Set ImgBack.Picture = New_BackImage
    PropertyChanged "BackImage"
End Property

Public Property Get BackGroundAlignment() As BackGroundAlign
    BackGroundAlignment = m_PicAlignment
End Property

Public Property Let BackGroundAlignment(ByVal New_Align As BackGroundAlign)
    m_PicAlignment = New_Align
    PropertyChanged "BackGroundAlignment"
    Call RenderLB
End Property

Public Property Get FullRowSelect() As Boolean
    FullRowSelect = m_FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_Row As Boolean)
    m_FullRowSelect = New_Row
    PropertyChanged "BackGroundAlignment"
    Call RenderLB
End Property

Public Property Get GridLines() As Boolean
    GridLines = m_GridLines
End Property

Public Property Let GridLines(ByVal vNewGrid As Boolean)
    m_GridLines = vNewGrid
    PropertyChanged "GridLines"
    Call RenderLB
End Property

Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = PicBase.hdc
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = PicBase.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    PicBase.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Private Sub PicBase_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub PicBase_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub PicBase_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub PicBase_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = PicBase.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set PicBase.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = PicBase.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    PicBase.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Public Property Get Sorted() As SortOrder
    Sorted = m_sort
End Property

Public Property Let Sorted(ByVal New_Sort As SortOrder)
    m_sort = New_Sort
    Call Sort(New_Sort)
    Call RenderLB
    PropertyChanged "Sorted"
End Property

Public Property Get Hint() As String
Attribute Hint.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    Hint = PicBase.ToolTipText
End Property

Public Property Let Hint(ByVal New_ToolTipText As String)
    PicBase.ToolTipText = New_ToolTipText
    PropertyChanged "Hint"
End Property

Public Property Get BorderStyle() As bStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As bStyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get IconMaskColor() As OLE_COLOR
Attribute IconMaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    IconMaskColor = m_IconMaskColor
End Property

Public Property Let IconMaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_IconMaskColor = New_MaskColor
    PropertyChanged "IconMaskColor"
End Property

Public Property Get IconStrip() As Picture
Attribute IconStrip.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set IconStrip = PicIcons.Picture
End Property

Public Property Set IconStrip(ByVal New_IconStrip As Picture)
    Set PicIcons.Picture = New_IconStrip
    PropertyChanged "IconStrip"
    Call RenderLB
End Property

Public Property Get ShowIcons() As Boolean
    ShowIcons = m_ShowIcons
End Property

Public Property Let ShowIcons(ByVal vNewShow As Boolean)
    m_ShowIcons = vNewShow
    Call RenderLB
    PropertyChanged "ShowIcons"
End Property

Public Property Get ImageHeight() As Integer
    ImageHeight = m_ImageHeight
End Property

Public Property Let ImageHeight(ByVal vNewHeight As Integer)
    m_ImageHeight = vNewHeight
    PropertyChanged "ImageHeight"
End Property

Public Property Get ImageWidth() As Integer
    ImageWidth = m_ImageWidth
End Property

Public Property Let ImageWidth(ByVal vNewWidth As Integer)
    m_ImageWidth = vNewWidth
    PropertyChanged "ImageHeight"
End Property

Public Property Get TextAlignment() As TextAlign
    TextAlignment = m_TextAlignment
End Property

Public Property Let TextAlignment(ByVal vNewAlignment As TextAlign)
    m_TextAlignment = vNewAlignment
    PropertyChanged "TextAlignment"
    Call RenderLB
End Property

Public Property Get gStartColor() As OLE_COLOR
    gStartColor = m_StartColor
End Property

Public Property Let gStartColor(ByVal vNewSC As OLE_COLOR)
    m_StartColor = vNewSC
    PropertyChanged "gStartColor"
    Call RenderLB
End Property

Public Property Get gEndColor() As OLE_COLOR
    gEndColor = m_EndColor
End Property

Public Property Let gEndColor(ByVal vNewEC As OLE_COLOR)
    m_EndColor = vNewEC
    PropertyChanged "gEndColor"
    Call RenderLB
End Property

Public Property Get gDirection() As GRADIENT_DIR
    gDirection = m_Direction
End Property

Public Property Let gDirection(ByVal vNewGDir As GRADIENT_DIR)
    m_Direction = vNewGDir
    PropertyChanged "gDirection"
    Call RenderLB
End Property
