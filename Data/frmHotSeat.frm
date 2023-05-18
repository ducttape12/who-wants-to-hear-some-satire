VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmHotSeat 
   BorderStyle     =   0  'None
   Caption         =   "Who Wants to Hear Some Satire?"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "frmHotSeat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmHotSeat.frx":030A
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MCI.MMControl mmcStart 
      Height          =   495
      Left            =   9120
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
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Let's &Play!"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Level:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   6000
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label lblUserName 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmHotSeat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' User wants to start the next question
Private Sub cmdStart_Click()
    ' Temporarly disable this button to prevent reclicks
    cmdStart.Enabled = False
    
    ' Play start sound
    mmcStart.FileName = "Start.wav"
    mmcStart.Command = "Load"
    mmcStart.Command = "Open"
    mmcStart.Command = "Play"

End Sub
' When the sub has focus, do this
Private Sub Form_Activate()
    ' Win
    ' ---
    ' If the final question has been asked, close this form and go to then end
    If (frmQuestion.intCompleted = 10) Then
        ' Open the final window
        frmComplete.Show
        ' Put the user' name on the check
        frmComplete.lblUserName.Caption = lblUserName.Caption
        
        ' Close the open windows
        Unload frmHotSeat
        Unload Me
        Unload frmCounter
        
    End If



    ' Set the current level label to the current level
    ' ------------------------------------------------
    ' Set the default part of lblLevel caption
    lblLevel.Caption = "Current Level:" + Chr$(13) + "$"
    
    ' Set the current question level
    Select Case frmQuestion.intCompleted
        Case 0
            lblLevel.Caption = lblLevel.Caption + "500"
            
        Case 1
            lblLevel.Caption = lblLevel.Caption + "1,000"
            
        Case 2
            lblLevel.Caption = lblLevel.Caption + "4,000"
            
        Case 3
            lblLevel.Caption = lblLevel.Caption + "8,000"
            
        Case 4
            lblLevel.Caption = lblLevel.Caption + "32,000"
            
        Case 5
            lblLevel.Caption = lblLevel.Caption + "64,000"
            
        Case 6
            lblLevel.Caption = lblLevel.Caption + "125,000"
            
        Case 7
            lblLevel.Caption = lblLevel.Caption + "250,000"
            
        Case 8
            lblLevel.Caption = lblLevel.Caption + "500,000"
            
        Case 9
            lblLevel.Caption = lblLevel.Caption + "1 Million"
            
    End Select


End Sub
' Get the question form going, but don't do anything with it
Private Sub Form_Load()
    ' Open form
    frmQuestion.Show
        
    ' Hide it right away
    frmQuestion.Hide
End Sub

' Song is done playing, so hide this window and start the next one
Private Sub mmcStart_Done(NotifyCode As Integer)
    ' Close this song (may not return to this screen if this is the final question)
    mmcStart.Command = "Close"
    
    ' Reenable use of the go button
    cmdStart.Enabled = True
    
    ' Open the question form
    frmQuestion.Show
    ' Set the activated button to be Answer A
    frmQuestion.optA.Value = True
    
    
    ' Launch the question getting sub on the Question form
    frmQuestion.GetNewQuestion
    
    
    ' Get music playing on Question form
    ' ----------------------------------
    ' Find what area in, and set music for that
    
    ' Easy round
    If (frmQuestion.intCompleted <= 2) Then
        frmQuestion.mmcMusic.FileName = "Round1.mid"
    End If
    
    ' Medium round
    If ((frmQuestion.intCompleted > 2) And (frmQuestion.intCompleted <= 5)) Then
        frmQuestion.mmcMusic.FileName = "Round2.mid"
    End If
    
    ' Hard round
    If ((frmQuestion.intCompleted > 5) And (frmQuestion.intCompleted <= 9)) Then
        frmQuestion.mmcMusic.FileName = "Round3.mid"
    End If
    
    ' 1 million
    If (frmQuestion.intCompleted = 10) Then
        frmQuestion.mmcMusic.FileName = "Final.mid"
    End If
    
    ' Start playing music
    frmQuestion.mmcMusic.Command = "Open"
    frmQuestion.mmcMusic.Command = "Load"
    frmQuestion.mmcMusic.Command = "Play"
    
    
       
    ' Make sure all answers are enabled (may be disabled if user used 50-50 last turn)
    frmQuestion.optA.Enabled = True
    frmQuestion.optB.Enabled = True
    frmQuestion.optC.Enabled = True
    frmQuestion.optD.Enabled = True
    
    ' Hide this window
    Me.Hide

End Sub
