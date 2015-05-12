Attribute VB_Name = "modProcess"
Option Explicit


' Very important piece of the program
' this recurring sub continuously calls itself
' and then call the Excel request handler object to do the work.
Sub Loop_Sub_Refresh()
    
    ' if users wants to stops the app, close
    If (CONTINUE_loop = False) And (RUNNING_loop = True) Then
        MsgBox "Link Stopped"
        RUNNING_loop = False
        Exit Sub
    End If
    
    ' if app has gone down: warn the user, then close
    If (CONTINUE_loop = False) And (RUNNING_loop = False) Then
        MsgBox "Link is closed," & vbCrLf & "use <Open Link> to activate."
        Exit Sub
    End If
    
    ' if we are here, the app is running
    RUNNING_loop = True
    
    'FOREVER
    Application.OnTime Now + TimeValue("00:00:01"), "Loop_Sub_Refresh"

    ' check that app is up
    Call checkApp
    
    ' finally, DO THE WORK
    If Not oApp Is Nothing Then
        ' handle all the pending requests
        Call oApp.cycleRequests
    End If

End Sub






