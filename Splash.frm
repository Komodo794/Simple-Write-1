VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Splash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6045
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Index           =   6700
      Left            =   3840
      Top             =   2640
   End
   Begin VB.Timer Timer4 
      Interval        =   5000
      Left            =   3000
      Top             =   2520
   End
   Begin VB.Timer Timer3 
      Interval        =   3000
      Left            =   2280
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Interval        =   1500
      Left            =   1440
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   3333
      Left            =   4920
      Top             =   360
   End
   Begin VB.Label Creator 
      Alignment       =   2  'Center
      Caption         =   "@2024 | Created by LukeAtlantis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Simple Write 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   5055
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


Private Sub Timer1_Timer()
Splash.Visible = False
Splash.Enabled = False
Main.Visible = True
Exit Sub
End Sub
