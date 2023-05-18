Attribute VB_Name = "InternetTimerSubs"
Option Explicit
' Sub for ending the program
' **************************
'
' Unloads all open forms
' Public so can be called by forms
Public Sub Quit()
    ' For unloading all open forms
    Dim intCount As Integer

    ' Unload all forms
    For intCount = (Forms.Count - 1) To 0 Step -1
        Unload Forms(intCount)
    Next intCount
    
    ' Quit the program
    End

End Sub
