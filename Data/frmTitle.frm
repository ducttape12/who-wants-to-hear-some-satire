VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmTitle 
   BorderStyle     =   0  'None
   Caption         =   "Who Wants to Hear Some Satire?"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "frmTitle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTitle.frx":030A
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl mmcStart 
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   6600
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
   Begin MCI.MMControl mmcBGMusic 
      Height          =   495
      Left            =   9120
      TabIndex        =   1
      Top             =   6600
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
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   4193
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' User has clicked "Start", so being the game
Private Sub cmdStart_Click()
    ' Variables
    ' ---------
    Dim strName As String       ' User's name
    
    
    ' Hide this button so can't be launched again
    cmdStart.Visible = False
    
    
    ' Get user's name
    ' ---------------
    strName = InputBox("Please enter your name:", "Who Wants to Hear Some Satire?", "Morris Chestnut")

    
    ' If user doesn't want to play, end the program
    If (strName = "") Then
        ' End program
        Quit
    End If
    
    ' Launch the inbetween screen
    frmHotSeat.Show
    ' Set user's name on the screen and on the question screen
    frmHotSeat.lblUserName.Caption = strName
    frmQuestion.lblUserName.Caption = strName
    ' Hide the form
    frmHotSeat.Hide
    
    
    ' Get ready
    ' ---------
    ' Stop music currently playing
    mmcBGMusic.Command = "Stop"
    mmcBGMusic.Command = "Close"
    
    ' Start gettting ready music playing
    mmcStart.FileName = "Play.mid"
    mmcStart.Command = "Open"
    mmcStart.Command = "Load"
    mmcStart.Command = "Play"
    
    ' Game is started when music is done playing
    
End Sub

' Perform these actions upon loading the form
Private Sub Form_Load()
    ' Make the audio player start playing
    mmcBGMusic.FileName = "Theme.MID"
    mmcBGMusic.Command = "Load"
    mmcBGMusic.Command = "Open"
    mmcBGMusic.Command = "Play"

End Sub
' Close down this screen
Private Sub Form_Unload(Cancel As Integer)
    ' Close the playing song
    mmcBGMusic.Command = "Stop"
    mmcBGMusic.Command = "Close"

End Sub

' Restart music playing
Private Sub mmcBGMusic_Done(NotifyCode As Integer)
    ' Rewind music
    mmcBGMusic.Command = "Stop"
    mmcBGMusic.Command = "Prev"
    mmcBGMusic.Command = "Play"

End Sub
' Music is done playing, so start game
Private Sub mmcStart_Done(NotifyCode As Integer)
    ' Close music
    mmcStart.Command = "Close"
    
    ' Close this screen
    Unload Me
    
    ' Show the hot seat
    frmHotSeat.Show

End Sub
