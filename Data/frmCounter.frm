VERSION 5.00
Begin VB.Form frmCounter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading..."
   ClientHeight    =   210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1425
   Icon            =   "frmCounter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleWidth      =   1425
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr1Min 
      Interval        =   60000
      Left            =   480
      Top             =   480
   End
   Begin VB.Label lblLoading 
      Alignment       =   2  'Center
      Caption         =   "Loading..."
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Do this action on loading
Private Sub Form_Load()
    ' Hide this form
    Me.Hide
    
    ' Show the title screen
    frmTitle.Show
End Sub
' This timer calls itself every 1 minute
Private Sub tmr1Min_Timer()
    ' Variables
    ' ---------
    Static intMinutes As Integer        ' The total number of minutes passed
    
    ' Error handler
    On Error GoTo ErrorHandler
    
    ' Add one to the total number of minutes passed
    intMinutes = intMinutes + 1
    
    
    ' If 8 minutes have passed and not last question, set to the last question
    If ((intMinutes >= 8) And (frmQuestion.intCompleted < 9)) Then
        frmQuestion.intCompleted = 9
    End If
    
    ' End sub
    Exit Sub
    
    
' Error handler
ErrorHandler:
' The only error that could most likely occur is the Question form not being open, so
' simply just skip this sub and try again in one minute

End Sub
