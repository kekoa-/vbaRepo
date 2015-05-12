VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------
' ## BASED ON:
' Module    : cAppEvents
' Company   : JKP Application Development Services (c) 2008
' Author    : Jan Karel Pieterse
' Created   : 2-6-2008
' Reviewe   : 2015-02-02 DB
' Purpose   : Handles Excel Application events
' Purpose   : Bloomberg-Like
'-------------------------------------------------------------------------
Option Explicit

'This object variable will hold the object who's events we want to respond to
'Note the "WithEvents" keyword, which is what we need to tell VBA it is an object
'with events.
Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Public cRequestList As Collection

Public nIter As Long
Public nSeconds_Loop As Long
'



' MAIN method of this object,
' execute the requests
Public Sub cycleRequests()

    ' use this helper object!
    Dim oH As cRequestHandler
    Set oH = New cRequestHandler
    
    Call oH.handleRequestList(cRequestList)
    
End Sub


Private Sub App_SheetCalculate(ByVal sh As Object)
    Call checkApp
    Call checkRun
End Sub



Private Sub App_WorkbookOpen(ByVal wb As Workbook)
End Sub

Private Sub Class_Terminate()
    Set App = Nothing
End Sub




