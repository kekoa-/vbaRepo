VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' in this request class, are stored the calls details.
' any getTable() function call from Excel generate a Requset object, that is then handled by a
' RequestHandler object



' ref to the caller workbook, not used atm
Public wb_caller As Workbook
' ref to the caller worksheet, not used atm
Public ws_caller As Worksheet

' ref to the caller range
Public rg_caller As Range
' ref to the log range
Public rg_run_log As Range
' ref to the run details range
Public rg_run_details As Range
' ref to the output range
Public rg_target As Range


' output row, not used atm
Public row_target As Long
' output column, not used atm
Public col_target As Long
' output rows and columns count
Public row_target_count As Long, col_target_count As Long

' request string
Public request As String
Public datatype As String
Public key As String
Public param1 As String
Public param2 As String
Public param3 As String

' tell if the request has already been handled
Public handled As Boolean

Public outputCleaned As Boolean


Public highlighted As Boolean
Public highlighted_count As Integer

' output by row?
Public output_byRow As Boolean
'
'


' check highlight function, if has been highlighted then undo highlight
Public Sub checkHighlight()
    If highlighted = True Then
        highlighted_count = highlighted_count + 1
        If highlighted_count >= 2 Then
            Call undoHighlight
        End If
    End If
End Sub

' highlights the caller range
' AFTER putting data in the target range
' GREEN color
Public Sub doHighlight()
    highlighted = True
    highlighted_count = 0
    If Not rg_caller Is Nothing Then
        rg_caller.Interior.Color = RGB(0, 255, 0)
    End If
End Sub

' undo the highlight of the caller range
Public Sub undoHighlight()
    highlighted = False
    highlighted_count = 0
    If Not rg_caller Is Nothing Then
        rg_caller.Interior.ColorIndex = xlNone
    End If
    Set rg_caller = Nothing
End Sub


' cleans the target range
Public Sub cleanOutput()
    Dim ws As Worksheet
    Dim rg_clean As Range
    
    ' highlights the caller range
    ' when clearing the target range
    ' LIGHT BLUE color
    If Not rg_caller Is Nothing Then
        rg_caller.Interior.Color = RGB(0, 255, 255)
    End If
    
    ' does XL_DOWN to clean the existing data-- user should be careful of existing data
    If Not rg_target Is Nothing Then
        Set ws = rg_target.Parent
        Set rg_clean = ws.Cells.Range(rg_target.Cells(1, 1), rg_target.Cells(1, 1).End(xlDown).Cells(1, col_target_count))
        rg_clean.ClearContents
        Set rg_clean = Nothing
        Set ws = Nothing
    End If
    
    outputCleaned = True

End Sub

' init the REQUEST object
Public Sub init()
    ' not yet handled!
    handled = False
    highlighted = False
    outputCleaned = False
    row_target_count = 1000
    col_target_count = 1
    output_byRow = False
End Sub

' destroys the request object
Public Sub destruct()
    Set wb_caller = Nothing
    Set ws_caller = Nothing
    ' this is cleaned in another point
'    Set rg_caller = Nothing
    Set rg_run_details = Nothing
    Set rg_target = Nothing
End Sub




