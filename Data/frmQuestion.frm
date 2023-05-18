VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmQuestion 
   BorderStyle     =   0  'None
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   495
   ClientWidth     =   9600
   Icon            =   "frmQuestion.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmQuestion.frx":030A
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuitProgram 
      Caption         =   "X"
      Height          =   255
      Left            =   9360
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Quit"
      Top             =   0
      Width           =   255
   End
   Begin MCI.MMControl mmcMusic 
      Height          =   495
      Left            =   9120
      TabIndex        =   11
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
   Begin MCI.MMControl mmcPhone 
      Height          =   495
      Left            =   120
      TabIndex        =   10
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
   Begin VB.CommandButton cmdAudience 
      Height          =   615
      Left            =   8160
      Picture         =   "frmQuestion.frx":E134E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ask the audience their opinion about the question."
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdPhone 
      Height          =   615
      Left            =   7200
      Picture         =   "frmQuestion.frx":E2A8A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Phone a friend for assistance"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmd5050 
      Height          =   615
      Left            =   6240
      Picture         =   "frmQuestion.frx":E4252
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Take away two wrong answers, leaving one incorrect and one correct answer."
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdFinalAnswer 
      Caption         =   "This is my final &answer!"
      Height          =   495
      Left            =   3533
      TabIndex        =   8
      Top             =   6600
      Width           =   2535
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H80000012&
      Caption         =   "Now Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   600
      Left            =   5730
      TabIndex        =   4
      Top             =   4620
      Width           =   3000
   End
   Begin VB.OptionButton optC 
      BackColor       =   &H80000012&
      Caption         =   "Now Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   600
      Left            =   1350
      TabIndex        =   3
      Top             =   4620
      Width           =   3000
   End
   Begin VB.OptionButton optB 
      BackColor       =   &H80000012&
      Caption         =   "Now Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   600
      Left            =   5730
      TabIndex        =   2
      Top             =   3840
      Width           =   3000
   End
   Begin VB.OptionButton optA 
      BackColor       =   &H80000012&
      Caption         =   "Now Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   600
      Left            =   1350
      TabIndex        =   1
      Top             =   3840
      Value           =   -1  'True
      Width           =   3000
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "UserName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   290
      Left            =   435
      TabIndex        =   9
      Top             =   290
      Width           =   1725
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Now Loading..."
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
      Height          =   1200
      Left            =   1200
      TabIndex        =   0
      Top             =   2280
      Width           =   7335
   End
End
Attribute VB_Name = "frmQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Global Variables
' ----------------
' Holds the real answer (can be accessed in other subs now)
Public intRealAnswer As Integer
Public intCompleted As Integer  ' Holds the number of questions completed (used
                                ' to step through dummy data)
Public bol5050Status As Boolean ' Whether cmd5050 is enabled or not (for use during
                                ' phone a friend
Public bolAudienceStatus As Boolean ' Whther cmdAudience is enabled or not


' Gets a new question for the form (public b/c launched from another form)
Public Sub GetNewQuestion()
    ' Variables
    ' ---------
    Dim strQuestion As String       ' The question from the file
    Dim strAnswerA As String        ' What to display as the first answer
    Dim strAnswerB As String        ' What to display as the second answer
    Dim strAnswerC As String        ' What to display as the third answer
    Dim strAnswerD As String        ' What to display as the forth answer
    Dim strDummy As String          ' To step through read data
    
    Dim intCount As Integer         ' For looping
    Dim intCount2 As Integer        ' For looping
        
    Dim intFreeNum As Integer       ' Holds a free number to work with files
                                    
    
    ' If an error occurs, end the program
    On Error GoTo ErrorHandler
    
    
    ' Open the file
    ' -------------
    ' Get a free number to work with
    intFreeNum = FreeFile
    
    ' Open the file for reading from
    Open "QA.txt" For Input As intFreeNum
    
    
    ' Read next questions from the file
    ' ---------------------------------
    ' First, step through all the used sets of questions
    For intCount = 0 To intCompleted Step 1
        ' Count through all the questions in each set
        For intCount2 = 1 To 6 Step 1
            ' Read dummy data
            Input #intFreeNum, strDummy
        Next
    Next
    
    ' Now, load the quetions into the form
    Input #intFreeNum, strQuestion
    Input #intFreeNum, strAnswerA
    Input #intFreeNum, strAnswerB
    Input #intFreeNum, strAnswerC
    Input #intFreeNum, strAnswerD
    Input #intFreeNum, intRealAnswer
    
    ' Close the file
    Close #intFreeNum
    
    ' Add to the total number of questions completed
    intCompleted = intCompleted + 1
    
    ' If the real answer is too high or low (<1 or >4) then end the program with an
    ' error
    If ((intRealAnswer < 1) Or (intRealAnswer > 4)) Then
        GoTo ErrorHandler
    End If
    
    
    
    ' Save the question/answers to the form
    ' -------------------------------------
    lblQuestion.Caption = strQuestion
    optA.Caption = strAnswerA
    optB.Caption = strAnswerB
    optC.Caption = strAnswerC
    optD.Caption = strAnswerD
    
    
    ' End of the sub
    Exit Sub
    
    
' Error handler
' -------------
' Simply warn user and quit program
ErrorHandler:
    ' Message box return for warning user
    Dim intMBR As Integer
    
    intMBR = MsgBox("This program has caused an illegal operation and will be sued by the RIAA..." + Chr$(13) + "Okay, not really, but it still will close." + Chr$(13) + Chr$(13) + "Please reinstall to fix this problem.", vbOKOnly + vbCritical, "Oops...")
    
    Quit
End Sub
' Get rid of two answers and leave two remaining
Private Sub cmd5050_Click()
    ' Variables
    ' ---------
    Dim intRemaining As Integer     ' The incorrect answer that will remain

    ' Start randomization
    ' -------------------
    Randomize Timer

    ' Disable use of 50-50 again
    ' --------------------------
    cmd5050.Enabled = False
    
    
    
    ' Reduce answers
    ' --------------
    ' Loop until the remaining answer isn't the real answer
    Do
        ' Get a random number 1-4
        intRemaining = Int(4 * Rnd(1)) + 1
        
    ' Loop until the remaining doesn't equal the real answer
    Loop While (intRemaining = intRealAnswer)
    
    
    ' Get rid of two answers
    ' ----------------------
    ' First, make all the answers unavaiblable
    optA.Enabled = False
    optB.Enabled = False
    optC.Enabled = False
    optD.Enabled = False


    ' Now, find and enable the real answer
    Select Case intRealAnswer
        ' A
        Case 1
            optA.Enabled = True
            
        ' B
        Case 2
            optB.Enabled = True
            
        ' C
        Case 3
            optC.Enabled = True
            
        ' D
        Case 4
            optD.Enabled = True
            
    End Select
    
    ' Next, find the remaining answer and enable it
    Select Case intRemaining
        ' A
        Case 1
            optA.Enabled = True
            
        ' B
        Case 2
            optB.Enabled = True
            
        ' C
        Case 3
            optC.Enabled = True
            
        ' D
        Case 4
            optD.Enabled = True
            
    End Select
    
End Sub
' Ask the audience for assistance
Private Sub cmdAudience_Click()
    ' Disable the Audience button
    cmdAudience.Enabled = False
    
    ' Disable use of the Final Answer button temporarly
    cmdFinalAnswer.Enabled = False
    
    
    ' Stop the main music playing
    mmcMusic.Command = "Stop"
    ' First, show the form
    frmAudience.Show
    
    ' Now, load the approate image into the form
    Select Case intRealAnswer
        ' A
        Case 1
            frmAudience.imgOpinion.Picture = LoadPicture("AudienceA.bmp")
            
        ' B
        Case 2
            frmAudience.imgOpinion.Picture = LoadPicture("AudienceB.bmp")
        
        ' C
        Case 3
            frmAudience.imgOpinion.Picture = LoadPicture("AudienceC.bmp")
        
        ' D
        Case 4
            frmAudience.imgOpinion.Picture = LoadPicture("AudienceD.bmp")
            
    End Select


End Sub

' User has choosen an answer, so see if it's correct
Private Sub cmdFinalAnswer_Click()
    ' Variables
    ' ---------
    Dim intMBR As Integer           ' The return from a message box
    Dim bolCorrect As Boolean       ' If user is correct or not
    
    
    ' Set default to user being incorrect
    bolCorrect = False
    
    
    ' Incorrect
    ' ---------
    ' See if real answer matches user's response
    Select Case intRealAnswer
        ' Correct answer is A
        Case 1
            ' If user is correct, make note - otherwise, make note that isn't correct
            If optA.Value = True Then
                bolCorrect = True
            End If
            
        ' Correct answer is B
        Case 2
            If optB.Value = True Then
                bolCorrect = True
            End If
            
        ' Correct answer is C
        Case 3
            If optC.Value = True Then
                bolCorrect = True
            End If
            
        ' Correct answer is D
        Case 4
            If optD.Value = True Then
                bolCorrect = True
            End If
            
    End Select
    
    ' If user isn't correct, tell them and let them retry
    If (bolCorrect = False) Then
        intMBR = MsgBox("Um... are you really sure?  Really, really sure?  (Hint hint)", vbOKOnly, "You're stupid!")
    End If
    
    
    ' Correct - end of this form
    ' --------------------------
    ' If user is correct, tell them and hide this form
    If (bolCorrect = True) Then
        intMBR = MsgBox("Ah geez... I'm really sorry.... but YOU JUST GOT IT RIGHT!", vbOKOnly, "Correct!")
        
        ' Stop and unload music from playing
        mmcMusic.Command = "Stop"
        mmcMusic.Command = "Close"
        
        ' Hide this form
        Me.Hide
        ' Show the hot seat form
        frmHotSeat.Show
               
    End If
    
End Sub
' Play phone a friend
Private Sub cmdPhone_Click()
    ' Disable user from doing anything but closing
    ' --------------------------------------------
    ' Save whether or not 50-50 and audience is enabled or not (so can disbale and,
    ' if needed, reenable later)
    bol5050Status = cmd5050.Enabled
    bolAudienceStatus = cmdAudience.Enabled
    
    
    ' Disable all buttons on form
    ' ---------------------------
    ' Disable use of this button again
    cmdPhone.Enabled = False
    
    ' Disable use of Final Answer until phone a friend is done playing
    cmdFinalAnswer.Enabled = False
    
    ' Other lifelines
    cmd5050.Enabled = False
    cmdAudience.Enabled = False
    
    ' Load and play phone a friend
    mmcPhone.FileName = "Phone.wav"
    mmcPhone.Command = "Open"
    mmcPhone.Command = "Load"
    mmcPhone.Command = "Play"

End Sub
' Make sure the user really wants to quit
Private Sub cmdQuitProgram_Click()
    ' Variables
    ' ---------
    Dim intMBR As Integer           ' The MsgBox return
    
    ' Ask user
    intMBR = MsgBox("Do you really want to quit?", vbYesNo + vbDefaultButton2 + vbQuestion, "Really quit?")
    
    ' If user wants to quit, then end the program
    If (intMBR = vbYes) Then
        ' Unload all music playing
        mmcMusic.Command = "Stop"
        mmcMusic.Command = "Close"
        ' Make sure the phone a friend is closed, too - if it isn't open, nothing will
        ' happen, otherwise
        mmcPhone.Command = "Stop"
        mmcPhone.Command = "Close"
        
        ' Quit the program
        Quit
    End If

End Sub

' Perform these actions on load
Private Sub Form_Load()
    ' Set the user's name
    lblUserName.Caption = frmHotSeat.lblUserName.Caption
    
End Sub
' Jump to the final question if not already there
Private Sub lblUserName_Click()
    ' Jump to the last question if it isn't the last question already
    If (intCompleted < 9) Then
        frmQuestion.intCompleted = 9
    End If

End Sub

' Restart music playing
Private Sub mmcMusic_Done(NotifyCode As Integer)
    mmcMusic.Command = "Prev"
    mmcMusic.Command = "Play"

End Sub

' Done playing, so end converstation
Private Sub mmcPhone_Done(NotifyCode As Integer)
    ' Close file
    mmcPhone.Command = "Close"

    ' Re-enable use of buttons if they should be enabled
    ' --------------------------------------------------
    ' Final answer button should always be reenabled
    cmdFinalAnswer.Enabled = True
    
    ' If the other life lines were originally enabled, enable them now
    cmd5050.Enabled = bol5050Status
    cmdAudience.Enabled = bolAudienceStatus
    
End Sub
