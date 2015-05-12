Attribute VB_Name = "modFixLinks"


Option Explicit


' Fix Links button,
' removes the external reference to the addin

Public Sub fixLinks()

    Dim wb As Workbook
    Dim sh As Worksheet
    
    Set wb = ActiveWorkbook
    
    For Each sh In wb.Worksheets
        sh.Cells.Replace What:="'K:\DB_Connect.xla'!", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Next
    

End Sub
