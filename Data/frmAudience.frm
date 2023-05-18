VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmAudience 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audience's Opinion"
   ClientHeight    =   5775
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6720
   Icon            =   "frmAudience.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl mmcMusic 
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   5280
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
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2745
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Image imgOpinion 
      Height          =   5145
      Left            =   0
      Picture         =   "frmAudience.frx":030A
      Top             =   0
      Width           =   6705
   End
End
Attribute VB_Name = "frmAudience"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Close this form
Private Sub cmdClose_Click()
    ' Close this form
    Unload Me
    
End Sub
' Do not let user click off of this form
Private Sub Form_Deactivate()
    ' Warn user they did an invalid thing
    Beep
    ' Set focus back to this form
    frmAudience.SetFocus

End Sub

' Do this on load
Private Sub Form_Load()
    ' Start background music playing
    mmcMusic.FileName = "Audience.mid"
    mmcMusic.Command = "Open"
    mmcMusic.Command = "Load"
    mmcMusic.Command = "Play"

End Sub
' Do these actions when the form closes
Private Sub Form_Unload(Cancel As Integer)
    ' Stop music playing here
    mmcMusic.Command = "Stop"
    mmcMusic.Command = "Close"
    
    ' Start music playing on question form
    frmQuestion.mmcMusic.Command = "Play"
    
    ' Re-enable use of the Final Answer button
    frmQuestion.cmdFinalAnswer.Enabled = True

End Sub

' Restart music playing
Private Sub mmcMusic_Done(NotifyCode As Integer)
    mmcMusic.Command = "Prev"
    mmcMusic.Command = "Play"

End Sub
