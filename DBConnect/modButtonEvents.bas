Attribute VB_Name = "modButtonEvents"

' open link button
' inits the app
Public Sub openLink_fromButton()

    Call SaveSetting(APPNAME, SECTION, "enabled", True)
    Call checkSettings
    
    If oApp Is Nothing Then
        MsgBox "Link Open"
        Application.Calculate
        Call checkApp
        Application.Calculate
        Call checkApp
    Else
        MsgBox "Link is open"
        Application.Calculate
        Call Loop_Sub_Refresh
    End If
End Sub


' refresh data button
' check that the app is initiated and running,
' then forces one full worksheet recalculation
Public Sub refreshData_fromButton()
    
    Call SaveSetting(APPNAME, SECTION, "enabled", True)
    Call checkSettings
    
    ' check that app is up
    Call checkApp
    ' check that app is running
    Call checkRun
    
    ' this would recompute all the open workbooks:
    ' Application.CalculateFull

    ' this recalculates the active sheet only
    ActiveWorkbook.ActiveSheet.Cells.Dirty
    Application.Calculate

End Sub


' close link button
' kills the app
Public Sub closeLink_fromButton()
    
    Call SaveSetting(APPNAME, SECTION, "enabled", False)

    CONTINUE_loop = False
    Set oApp = Nothing
    MsgBox "Link Closed"
End Sub

' "show version" button
Public Sub showVersion()
    Call MsgBox("Build version: 1.0.1", vbInformation, "DB Connect")
End Sub



