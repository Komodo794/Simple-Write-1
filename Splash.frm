VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Splash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   6360
      Top             =   960
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Progress As Integer

Function ProgressBar(Amount As Integer)
ProgressBar = Amount
End Function


Private Sub Image1_Click()

End Sub

Private Sub Timer1_Timer()
Splash.Visible = False
Splash.Enabled = False
Main.Visible = True
Exit Sub
End Sub

