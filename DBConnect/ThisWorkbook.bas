VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'-------------------------------------------------------------------------
' %%% BASED on
' Company   : JKP Application Development Services (c) 2005
' Author    : Jan Karel Pieterse
' Created   : 02-06-2008
' Purpose   : Workbook event code
' Update    : 2015-02-02 DB
'-------------------------------------------------------------------------
Option Explicit

Private Sub Workbook_Open()
    '-------------------------------------------------------------------------
    ' Procedure : Workbook_Open Created by Jan Karel Pieterse
    ' Company   : JKP Application Development Services (c) 2005
    '%%% BASED on
    ' Author    : Jan Karel Pieterse
    ' Created   : 12-12-2005
    ' Update    : 2015-02-02 DB
    ' Purpose   : Code run at opening of workbook
    '-------------------------------------------------------------------------
    
    'code inthis section enables manual calculation
    Dim tempWkbk As Workbook
    Application.EnableEvents = True
    Set tempWkbk = Workbooks.Add
    Application.EnableEvents = True
    Application.Calculation = xlManual
    Application.CalculateBeforeSave = False
    tempWkbk.Close savechanges:=False
    
    'Initialise the application
    ' add menus
    Call AddMenus
    ' check app settings
    Call checkSettings
    ' init the app objects
    Call checkApp
    ' start the program
    Call checkRun

End Sub
