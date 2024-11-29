VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Main 
   Caption         =   "Simple Write 1"
   ClientHeight    =   3525
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox SizeBox 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Text            =   "5"
      Top             =   50
      Width           =   350
   End
   Begin VB.TextBox FontBox 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Font"
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
      TextRTF         =   $"Main.frx":0000
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
      Begin VB.Menu menuFormat 
         Caption         =   "Format"
         Index           =   4
         Begin VB.Menu formatItalics 
            Caption         =   "Italic"
            Shortcut        =   ^I
         End
         Begin VB.Menu formatBold 
            Caption         =   "Bold"
            Shortcut        =   ^B
         End
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
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Text As String
Dim TextSize As Integer

Function UpdateFontBox(Font As String)
FontBox.Text = Font
End Function

Private Sub Form_Load()
RichTextBox1.Text = ""
FontBox.Text = RichTextBox1.Font.Name
FontBox.Font = RichTextBox1.Font
SizeBox.Text = RichTextBox1.Font.Size
End Sub

Private Sub Form_Resize()
RichTextBox1.Left = 0
RichTextBox1.Width = Me.Width - 115
RichTextBox1.Height = Me.Height * 0.99
End Sub

Private Sub RichTextBox1_Click()
SizeBox.Text = RichTextBox1.SelFontSize
FontBox.Text = RichTextBox1.SelFontName
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

Private Sub formatBold_Click()
If Not RichTextBox1.SelBold = True Then
RichTextBox1.SelBold = True
Else
RichTextBox1.SelBold = False
End If
End Sub

Private Sub formatItalics_Click()
If Not RichTextBox1.SelItalic = True Then
RichTextBox1.SelItalic = True
Else
RichTextBox1.SelItalic = False
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
