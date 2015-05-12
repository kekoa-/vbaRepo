Attribute VB_Name = "mod_Checks"
Option Explicit

Public Const DISP_NO_ERRORS = False
Public Const RANGE_EXISTS = True
Public Const RESULT_OK = True

' check that the worksheets exists and that the range exists
Public Function checkRangeExists(ByRef wb As Workbook, ByVal wsName As String, _
                            ByVal rngName As String, Optional ByVal dispErrors As Boolean = True) As Boolean
On Error GoTo errExit
        
    If Not (wb Is Nothing) Then
    If Not (wb.Worksheets(wsName) Is Nothing) Then
    If Not (wb.Worksheets(wsName).Range(rngName) Is Nothing) Then
        checkRangeExists = True
    End If
    End If
    End If
    
On Error GoTo 0
Exit Function
    
errExit:
    On Error GoTo 0
    If dispErrors = True Then
    ' display message
        MsgBox "Attenzione, non esiste il range:" & vbCrLf & _
            "Range Name: " & rngName & vbCrLf & _
            "Worksheet Name: " & wsName & vbCrLf & vbCrLf & _
            "Il campo corrispondente non viene aggiornato." & vbCrLf & vbCrLf & "<mod_Checks:checkRangeExists>"
    End If
    checkRangeExists = False
endFunction:
End Function

' check that the range contains data of the correct type
Public Function checkSingleRangeType(ByRef rg As Range, ByVal expectedType_ As String, _
                               Optional ByVal dispErrors As Boolean = True) As Boolean
    Dim cU As cUpdater
    Set cU = New cUpdater
    checkSingleRangeType = cU.checkSingleRangeType(rg, expectedType_, dispErrors)
End Function


' check if recordset is nonEmpty
Public Function check_Recordset_NotEmpty(ByRef rs As Recordset) As Boolean
On Error GoTo exitFalse
    If rs Is Nothing Then
        check_Recordset_NotEmpty = False
    Else
        rs.MoveFirst
        check_Recordset_NotEmpty = True
    End If

    On Error GoTo 0
    Exit Function
exitFalse:
    On Error GoTo 0
    check_Recordset_NotEmpty = False
End Function














