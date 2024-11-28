VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Main 
   Caption         =   "Simple Write 1"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5953
      _Version        =   393217
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
      End
      Begin VB.Menu menuFonts 
         Caption         =   "Fonts"
         Index           =   2
         Begin VB.Menu menuArial 
            Caption         =   "Arial"
            Index           =   3
         End
         Begin VB.Menu menuSerif 
            Caption         =   "MS Serif"
            Index           =   3
         End
      End
      Begin VB.Menu menuFormat 
         Caption         =   "Format"
         Index           =   4
         Begin VB.Menu formatRegular 
            Caption         =   "Regular"
         End
         Begin VB.Menu formatBold 
            Caption         =   "Bold"
         End
      End
      Begin VB.Menu menuSize 
         Caption         =   "Size"
         Begin VB.Menu size8 
            Caption         =   "Size 8"
         End
         Begin VB.Menu size10 
            Caption         =   "Size 10"
         End
         Begin VB.Menu size12 
            Caption         =   "Size 12"
         End
         Begin VB.Menu size14 
            Caption         =   "Size 14"
         End
         Begin VB.Menu size18 
            Caption         =   "Size 18"
         End
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Help"
      Index           =   3
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Text As String

Private Sub Form_Load()
RichTextBox1.Text = ""
End Sub

Private Sub Form_Resize()
RichTextBox1.Left = 0
RichTextBox1.Width = Me.Width - 115
RichTextBox1.Height = Me.Height * 0.99
End Sub

Private Sub menuHelp_Click(Index As Integer)
About.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub formatBold_Click()
RichTextBox1.SelBold = True
End Sub

Private Sub formatRegular_Click()
RichTextBox1.SelBold = False
RichTextBox1.SelItalic = False
End Sub

Private Sub size8_Click()
RichTextBox1.SelFontSize = 8
End Sub

Private Sub size10_Click()
RichTextBox1.SelFontSize = 10
End Sub

Private Sub size12_Click()
RichTextBox1.SelFontSize = 12
End Sub

Private Sub size14_Click()
RichTextBox1.SelFontSize = 14
End Sub

Private Sub size18_Click()
RichTextBox1.SelFontSize = 18
End Sub



Private Sub menuArial_Click(Index As Integer)
Text = RichTextBox1.Text
RichTextBox1.SelFontName = "Arial"
End Sub

Private Sub menuSerif_Click(Index As Integer)
Text = RichTextBox1.Text
RichTextBox1.SelFontName = "MS Sans Serif"
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
