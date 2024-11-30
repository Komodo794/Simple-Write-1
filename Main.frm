VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   Caption         =   "Simple Write 1"
   ClientHeight    =   3525
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7545
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton StrikethroughButton 
      Cancel          =   -1  'True
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Strikethrough"
      Top             =   50
      Width           =   315
   End
   Begin VB.CommandButton UnderLineButton 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Underline"
      Top             =   50
      Width           =   315
   End
   Begin VB.Frame TextColor 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      ToolTipText     =   "Font Color"
      Top             =   50
      Width           =   315
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton AllignLeft 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Allign Left"
      Top             =   50
      Width           =   315
   End
   Begin VB.CommandButton AllignMiddle 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Allign Center"
      Top             =   50
      Width           =   315
   End
   Begin VB.CommandButton AllignRight 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Allign Right"
      Top             =   50
      Width           =   315
   End
   Begin VB.CommandButton ItalicsButton 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Italics"
      Top             =   50
      Width           =   315
   End
   Begin VB.CommandButton BoldButton 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Bold"
      Top             =   50
      Width           =   315
   End
   Begin VB.TextBox SizeBox 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Text            =   "5"
      ToolTipText     =   "Font Size"
      Top             =   50
      Width           =   350
   End
   Begin VB.TextBox FontBox 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Font"
      ToolTipText     =   "Font"
      Top             =   50
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   400
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Main.frx":0000
   End
   Begin VB.OLE OLE1 
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
         Index           =   1
      End
   End
   Begin VB.Menu menuEdit 
      Caption         =   "Edit"
      Index           =   2
      Begin VB.Menu menuClear 
         Caption         =   "Clear"
         Index           =   2
         Shortcut        =   ^W
      End
      Begin VB.Menu menuUndo 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu menuBullet 
         Caption         =   "Bullet"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu fontsMenu 
      Caption         =   "Fonts"
      Visible         =   0   'False
      Begin VB.Menu fontsSerif 
         Caption         =   "MS Sans Serif"
      End
      Begin VB.Menu fontsArial 
         Caption         =   "Arial"
      End
      Begin VB.Menu fontsImpact 
         Caption         =   "Impact"
      End
      Begin VB.Menu fontsSystem 
         Caption         =   "System"
      End
      Begin VB.Menu fontsTimes 
         Caption         =   "Times New Roman"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Text As String
Dim TextSize As Integer
Dim DefaultButton As String
Dim DownButton As String
Dim PreviousText As String

Function UpdateFontBox(Font As String)
FontBox.Text = Font
End Function

Function IsValidSelection(Text As String)
IsValidSelection = False
If Text <> vbNullString Then
IsValidSelection = True
End If
End Function

Private Sub AllignLeft_Click(Index As Integer)
RichTextBox1.SelAlignment = 0
End Sub

Private Sub AllignMiddle_Click(Index As Integer)
RichTextBox1.SelAlignment = 2
End Sub

Private Sub AllignRight_Click(Index As Integer)
RichTextBox1.SelAlignment = 1
End Sub

Private Sub menuBullet_Click()
If Not RichTextBox1.SelBullet = True Then
RichTextBox1.SelBullet = True
Else
RichTextBox1.SelBullet = False
End If
End Sub

Private Sub TextColor_Click()
CommonDialog1.ShowColor
TextColor.BackColor = CommonDialog1.Color
RichTextBox1.SelColor = CommonDialog1.Color
End Sub

Private Sub Form_Load()
RichTextBox1.Text = ""
FontBox.Text = RichTextBox1.Font.Name
FontBox.Font = RichTextBox1.Font
SizeBox.Text = RichTextBox1.Font.Size
DefaultButton = &H8000000F
DownButton = &H80000016
End Sub

Private Sub Form_Resize()
RichTextBox1.Left = 0
RichTextBox1.Width = Me.Width - 115
RichTextBox1.Height = Main.Height - 1050
End Sub

Private Sub RichTextBox1_Click()
SizeBox.Text = RichTextBox1.SelFontSize
FontBox.Text = RichTextBox1.SelFontName
TextColor.BackColor = RichTextBox1.SelColor

If RichTextBox1.SelBold = True Then
BoldButton(0).BackColor = DownButton
Else
BoldButton(0).BackColor = DefaultButton
End If

If RichTextBox1.SelItalic = True Then
ItalicsButton(1).BackColor = DownButton
Else
ItalicsButton(1).BackColor = DefaultButton
End If

If RichTextBox1.SelUnderline = True Then
UnderLineButton(0).BackColor = DownButton
Else
UnderLineButton(0).BackColor = DefaultButton
End If

If RichTextBox1.SelStrikeThru = True Then
StrikethroughButton(1).BackColor = DownButton
Else
StrikethroughButton(1).BackColor = DefaultButton
End If

End Sub

Private Sub SizeBox_Change()
If Not SizeBox.Text = "" Then
RichTextBox1.SelFontSize = SizeBox.Text
End If
End Sub

Private Sub FontBox_Click()
PopupMenu fontsMenu
End Sub

Private Sub fontsArial_Click()
Text = RichTextBox1.Text
RichTextBox1.SelFontName = "Arial"
UpdateFontBox (RichTextBox1.SelFontName)
End Sub

Private Sub fontsSerif_Click()
Text = RichTextBox1.Text
RichTextBox1.SelFontName = "MS Sans Serif"
UpdateFontBox (RichTextBox1.SelFontName)
End Sub

Private Sub fontsImpact_Click()
Text = RichTextBox1.Text
RichTextBox1.SelFontName = "Impact"
UpdateFontBox (RichTextBox1.SelFontName)
End Sub

Private Sub fontsSystem_Click()
Text = RichTextBox1.Text
RichTextBox1.SelFontName = "System"
UpdateFontBox (RichTextBox1.SelFontName)
End Sub

Private Sub fontsTimes_Click()
Text = RichTextBox1.Text
RichTextBox1.SelFontName = "Times New Roman"
UpdateFontBox (RichTextBox1.SelFontName)
End Sub

Private Sub mnuHelp_Click()
About.Left = Me.Left
About.Top = Me.Top
About.Visible = True
End Sub

Private Sub menuExit_Click(Index As Integer)
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub BoldButton_Click(Index As Integer)
If Not RichTextBox1.SelBold = True Then
RichTextBox1.SelBold = True
BoldButton(0).BackColor = DownButton
Else
RichTextBox1.SelBold = False
BoldButton(0).BackColor = DefaultButton
End If
End Sub

Private Sub ItalicsButton_Click(Index As Integer)
If Not RichTextBox1.SelItalic = True Then
RichTextBox1.SelItalic = True
ItalicsButton(1).BackColor = DownButton
Else
RichTextBox1.SelItalic = False
ItalicsButton(1).BackColor = DefaultButton
End If
End Sub

Private Sub UnderLineButton_Click(Index As Integer)
If Not RichTextBox1.SelUnderline = True Then
RichTextBox1.SelUnderline = True
UnderLineButton(0).BackColor = DownButton
Else
RichTextBox1.SelUnderline = False
UnderLineButton(0).BackColor = DefaultButton
End If
End Sub

Private Sub StrikethroughButton_Click(Index As Integer)
If Not RichTextBox1.SelStrikeThru = True Then
StrikethroughButton(1).BackColor = DownButton
RichTextBox1.SelStrikeThru = True
Else
StrikethroughButton(1).BackColor = DefaultButton
RichTextBox1.SelStrikeThru = False
End If
End Sub


Private Sub menuClear_Click(Index As Integer)
Text = RichTextBox1.Text
If Not Text = "" Then
RichTextBox1.Text = ""
MsgBox "Text has been cleared!"
Else
MsgBox "No valid text was found!"
End If
End Sub
