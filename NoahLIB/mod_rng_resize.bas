Attribute VB_Name = "mod_rng_resize"
Option Explicit


' deletes the specified number of rows from the range
Public Function range_deleteRows(ByRef rg As Range, ByVal nRows As Long) As Boolean
    range_deleteRows = False
    On Error GoTo errFunction
    
    If (rg.Rows.count < nRows) Then
        MsgBox "Error [mod_rng_resize:range_deleteRows] : deleting too much rows"
        End
    End If
    
    Dim rg1, rg2, rg3 As Range
    Dim ws1 As Worksheet
    
    Set rg1 = rg.Cells(1, 1)
    
    Set rg2 = rg1.Cells(nRows, rg.Columns.count)
    
    Set ws1 = rg.Parent
    
    Set rg3 = ws1.Cells.Range(rg1, rg2)
    rg3.Select
    
    rg3.Delete Shift:=xlUp
    
    range_deleteRows = True
    On Error GoTo 0
    Exit Function
    
errFunction:
    On Error GoTo 0
    MsgBox "Error [mod_rng_resize:range_deleteRows] "
End Function


' inserts the specified number of rows into the range
Public Function range_addRows(ByRef rg As Range, ByVal nRows As Long) As Boolean
On Error GoTo errFunction
    range_addRows = False

    Dim rg1, rg2, rg3 As Range
    Dim ws1 As Worksheet
    
    Set rg1 = rg.Cells(rg.Rows.count, 1)
    
    Set rg2 = rg1.Cells(nRows, rg.Columns.count)
    
    Set ws1 = rg.Parent
    
    Set rg3 = ws1.Cells.Range(rg1, rg2)
    
    rg3.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    range_addRows = True
    On Error GoTo 0
    Exit Function
    
errFunction:
    On Error GoTo 0
    MsgBox "Error [mod_rng_resize:range_addRows] "
End Function



' resize the range to the specified row count
Public Function range_resizeRows(ByRef rg As Range, ByVal nRows As Long) As Boolean
On Error GoTo errFunction
    If rg Is Nothing Then Exit Function
    
    'not too small
    If nRows < 2 Then nRows = 2
    
    If rg.Rows.count < 2 Then
        MsgBox "Error [mod_rng_resize:resizeRangeRows] : range must have 2 or more rows"
        End
    End If
    
    If nRows > rg.Rows.count Then
        range_resizeRows = range_addRows(rg, nRows - rg.Rows.count)
    End If
    
    If nRows < rg.Rows.count Then
        range_resizeRows = range_deleteRows(rg, rg.Rows.count - nRows)
    End If
    
endFunction:
    range_resizeRows = True
    On Error GoTo 0
    Exit Function

errFunction:
    On Error GoTo 0
    MsgBox "Error [mod_rng_resize:resizeRangeRows] "
    
End Function



' resize the range to the specified row count, insertion of rows is performed on the worksheet
Public Function range_insertRows_worksheet(ByRef rg As Range, ByVal nRows As Long) As Boolean
    Dim nToAdd, i As Long
    Dim ws As Worksheet
    Dim rg2 As Range
    Dim strRow As String


    If nRows > rg.Rows.count Then
        nToAdd = nRows - rg.Rows.count
        Set ws = rg.Parent
        strRow = "" & rg.Row + 1 & ":" & rg.Row + 1
        Set rg2 = ws.Cells.Range(strRow)
        rg2.Select
        For i = 1 To nToAdd
            rg2.Insert Shift:=xlUp, CopyOrigin:=xlFormatFromLeftOrAbove
        Next
        range_insertRows_worksheet = True
    Else
        range_insertRows_worksheet = True
    End If
End Function












