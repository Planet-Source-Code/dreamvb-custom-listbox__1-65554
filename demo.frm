VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "ListBox Demo"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFullrow 
      Caption         =   "FullRow Select"
      Height          =   225
      Left            =   3120
      TabIndex        =   11
      Top             =   2895
      Width           =   1485
   End
   Begin VB.ComboBox cboAlign 
      Height          =   315
      Left            =   75
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   3705
      Width           =   2175
   End
   Begin VB.CheckBox chk2 
      Caption         =   "Show Grid Lines"
      Height          =   300
      Left            =   1485
      TabIndex        =   8
      Top             =   2895
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Back Styles"
      Height          =   1170
      Left            =   5760
      TabIndex        =   4
      Top             =   285
      Width           =   1425
      Begin VB.OptionButton Option3 
         Caption         =   "Gradient"
         Height          =   255
         Left            =   135
         TabIndex        =   7
         Top             =   795
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Texture"
         Height          =   255
         Left            =   135
         TabIndex        =   6
         Top             =   570
         Width           =   960
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Soild Color"
         Height          =   255
         Left            =   135
         TabIndex        =   5
         Top             =   300
         Width           =   1260
      End
   End
   Begin VB.CheckBox ch1 
      Caption         =   "Show Icons"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2895
      Width           =   1200
   End
   Begin VB.TextBox TxtInfo 
      Height          =   2340
      Left            =   3180
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   300
      Width           =   2460
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   660
      Left            =   6420
      TabIndex        =   1
      Top             =   4155
      Width           =   1605
   End
   Begin Project1.ListBoxFX ListBoxFX1 
      Height          =   2340
      Left            =   90
      TabIndex        =   0
      Top             =   315
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   4128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectionClolor =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FullRowSelect   =   0   'False
      ShowIcons       =   -1  'True
      gStartColor     =   16777215
      gEndColor       =   16744576
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Picture Alignment"
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   3465
      Width           =   1230
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Info()
    TxtInfo.Text = "ListCount = " & ListBoxFX1.ListCount & vbCrLf _
    & "Seleted Index = " & ListBoxFX1.ListIndex & vbCrLf _
    & "Selected Item = " & ListBoxFX1.ItemSelectedText & vbCrLf _
    & "Selected Item Key = " & ListBoxFX1.Key(ListBoxFX1.ListIndex)
End Sub
Sub LoadFonts()
Dim x As Integer, xImg As Integer

    For x = 0 To Screen.FontCount - 1
        ListBoxFX1.AddItem Screen.Fonts(x), "Key:" & x, xImg
        xImg = xImg + 1
        If xImg >= 47 Then xImg = 0
    Next x
End Sub

Private Sub cboAlign_Click()
    ListBoxFX1.BackGroundAlignment = cboAlign.ListIndex
End Sub

Private Sub ch1_Click()
    ListBoxFX1.ShowIcons = ch1
End Sub

Private Sub chk2_Click()
    ListBoxFX1.GridLines = chk2
End Sub

Private Sub Command1_Click()
ListBoxFX1.ShowIcons = True

End Sub

Private Sub Command2_Click()
ListBoxFX1.ShowIcons = False
End Sub

Private Sub chkFullrow_Click()
    ListBoxFX1.FullRowSelect = chkFullrow
End Sub

Private Sub cmdexit_Click()
    End
End Sub

Private Sub Form_Activate()
    LoadFonts
    ListBoxFX1.ListIndex = 0
    Info
    ch1_Click
End Sub

Private Sub Form_Load()
    Set ListBoxFX1.BackImage = LoadPicture(App.Path & "\bk.gif")
    Set ListBoxFX1.IconStrip = LoadPicture(App.Path & "\1.bmp")

    cboAlign.AddItem "TopLeft"
    cboAlign.AddItem "TopRight"
    cboAlign.AddItem "Center"
    cboAlign.AddItem "Tile"
    cboAlign.AddItem "BottomLeft"
    cboAlign.AddItem "BottomRight"
    cboAlign.ListIndex = 0
    
End Sub

Private Sub ListBoxFX1_Change()
    Info
End Sub

Private Sub ListBoxFX1_Click()
    Info
End Sub

Private Sub Option1_Click()
    ListBoxFX1.BackGroundStyle = bSoildColor
End Sub

Private Sub Option2_Click()
    ListBoxFX1.BackGroundStyle = bBitmap
End Sub

Private Sub Option3_Click()
    ListBoxFX1.BackGroundStyle = bGradient
End Sub
