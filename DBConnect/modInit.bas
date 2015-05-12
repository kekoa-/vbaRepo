Attribute VB_Name = "modInit"
Option Explicit

' Create a module level object variable that will keep the instance of the
' event listener in memory (and hence alive)
Public oApp As cAppEvents

' show debug can be set in the registry
Public SHOW_DEBUG As Boolean

' CONTINUE_loop : tells if user wants the app running
Public CONTINUE_loop As Boolean
' running_loop : tells the app is running already
Public RUNNING_loop As Boolean
' SHOW_logs can and SHOW_RUN_details be set in the registry
Public SHOW_logs As Boolean
Public SHOW_RUN_details As Boolean

' appname and section name in the registry
Public Const APPNAME = "Katarsis_Bloomberg"
Public Const SECTION = "Preferences"


' init App : instances the app objects
Sub InitApp()

    'Create a new instance of cAppEvents class
    Set oApp = New cAppEvents
    'Tell it to listen to Excel's events
    Set oApp.App = Application
    
    ' inits the request collection
    Set oApp.cRequestList = New Collection
    
    oApp.nIter = 0
    oApp.nSeconds_Loop = 2

End Sub

' check if app is up, if not, inits the app
Public Sub checkApp()
    If oApp Is Nothing Then
        Call InitApp
    End If
End Sub

' load app settings from registry
Public Sub checkSettings()
    CONTINUE_loop = GetSetting(APPNAME, SECTION, "enabled", False)
    SHOW_logs = GetSetting(APPNAME, SECTION, "showLogs", False)
    SHOW_RUN_details = GetSetting(APPNAME, SECTION, "showRunDetails", False)
    SHOW_DEBUG = GetSetting(APPNAME, SECTION, "showDebug", False)
End Sub

' check if app is running already.
' if app is not running but the user wants it running, runs the app
Public Sub checkRun()
    If (CONTINUE_loop = True) And (RUNNING_loop = False) Then
        RUNNING_loop = True
        Call Loop_Sub_Refresh
    End If
End Sub






