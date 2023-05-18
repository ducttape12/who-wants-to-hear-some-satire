VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmComplete 
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "frmComplete.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmComplete.frx":030A
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl mmcMusic 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   873
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "Sequencer"
      FileName        =   ""
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Thanks for Playing ""Who Wants to Hear Some Satire?"""
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   6600
      Width           =   9375
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2640
      Width           =   5415
   End
   Begin VB.Image imgCheck 
      Height          =   2760
      Left            =   1440
      Picture         =   "frmComplete.frx":E134E
      Top             =   1800
      Width           =   7065
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This has been a Keith Ott and Chris Connar Production"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   9375
   End
   Begin VB.Label lblCongrats 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations!  You've won $1 MILLION!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   9375
   End
End
Attribute VB_Name = "frmComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Close this open window
Private Sub cmdQuit_Click()
    ' Stop music playing
    mmcMusic.Command = "Stop"
    mmcMusic.Command = "Close"
    
    
    ' Close all open windows
    Quit

End Sub
' Perform these actions when the form loads
Private Sub Form_Load()
    ' Play winning music
    ' ------------------
    mmcMusic.FileName = "Credits.mid"
    mmcMusic.Command = "Open"
    mmcMusic.Command = "Load"
    mmcMusic.Command = "Play"

End Sub
' Reset music playing
Private Sub mmcMusic_Done(NotifyCode As Integer)
    mmcMusic.Command = "Prev"
    mmcMusic.Command = "Play"

End Sub
