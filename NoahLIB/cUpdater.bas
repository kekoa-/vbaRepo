VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public thisWb As Workbook


Public Function convertStringToInt(ByVal s As String) As Long
    convertStringToInt = 0
    If s = "y" Or s = "Y" Or s = "Yes" Or s = "TRUE" Or s = "True" Then convertStringToInt = 1
    'If s = "n" Or s = "N" Or s = "No" Or s = "FALSE" Or s = "" Then convertStringToInt = 0
End Function




' works on for SINGLE CELL ranges
' checks that the range contains data of the right type
' optionally displays errors or terminates program execution
Public Function checkSingleRangeType(ByRef rg As Range, _
                                    ByVal expectedType As String, _
                                    Optional ByVal dispErrors As Boolean = True, _
                                    Optional ByVal endOnError As Boolean = False) As Boolean
    
    checkSingleRangeType = False
    
    Dim matchedType As Boolean
    Dim s As String
    Dim d As Double
    Dim dat As Date
    
    If rg Is Nothing Then
        If dispErrors Then
            MsgBox "[cUpdater:checkSingleRangeType] : NULL Range", vbExclamation, "Error"
        End If
        If endOnError Then
            End
        End If
        Exit Function
    End If
    
    If IsError(rg) Then
        If dispErrors Then
            rg.Parent.Activate
            rg.Select
            MsgBox "[cUpdater:checkSingleRangeType] : Error found in this range ", vbExclamation, "Error"
        End If
        If endOnError Then
            End
        End If
        Exit Function
    End If

    matchedType = False
    
    
    On Error GoTo errFunctionType
    
    If expectedType = "STR" Then
        s = CStr(rg.value)
        checkSingleRangeType = True
        matchedType = True
    End If
    
    If expectedType = "DOUBLE" Then
        d = CDbl(rg.value)
        checkSingleRangeType = True
        matchedType = True
    End If
    
    If expectedType = "INT" Then
        d = CInt(rg.value)
        checkSingleRangeType = True
        matchedType = True
    End If
    
    If expectedType = "DATE" Then
        dat = CDate(rg.value)
        checkSingleRangeType = True
        matchedType = True
    End If
    
    If expectedType = "BOOL" Then
        s = rg.value
        If convertStringToInt(s) >= 0 Then
            checkSingleRangeType = True
            matchedType = True
        End If
    End If
    
    
    If Not matchedType Then
        rg.Parent.Activate
        rg.Select
        MsgBox "Error, type not matched: <" & expectedType & "> for range: " & rg.Name & vbCrLf & _
                "Module: cUpdater:checkSingleRangeType ", vbCritical, "Type Error"
        End
    End If
    
    Exit Function
    
errFunctionType:
    On Error GoTo 0
    If dispErrors Then
        rg.Parent.Activate
        rg.Select
        MsgBox "Warning, expected a <" & expectedType & "> value in this cell!" & vbCrLf & vbCrLf & "cUpdater:checkSingleRangeType"
        If endOnError Then
            End
        End If
    End If
    
End Function

' check that a multi-cell range contains data of the expected type
' iterates over SINGLE/MULTI CELL ranges
Public Function checkMultiRangeType(ByRef rg As Range, _
                                    ByVal expectedType As String, _
                                    Optional ByVal dispErrors As Boolean = True, _
                                    Optional ByVal endOnError As Boolean = False) As Boolean

Dim rIterator As Range
checkMultiRangeType = True
' cycle over cells in the range
For Each rIterator In rg.Cells
    If Not checkSingleRangeType(rIterator, expectedType, dispErrors, endOnError) Then
        checkMultiRangeType = False
    End If
Next

End Function

' helper function, get the range using the Link object
Public Function getRange(ByRef obj As cLinkField) As Range
    ' helper function
    On Error Resume Next
    Set getRange = Nothing
    Set getRange = thisWb.Worksheets(obj.WorksheetName).Range(obj.RangeName)
    On Error GoTo 0
End Function

' helper function, get the VALUE-range using the Link object
Public Function getRangeData(ByRef obj As cLinkField) As Range
    ' helper function
    On Error Resume Next
    Set getRangeData = Nothing
    Set getRangeData = thisWb.Worksheets(obj.WorksheetName).Range(obj.RangeName)
    On Error GoTo 0
End Function

' helper function, get the KEY-range using the Link object
Public Function getRangeKey(ByRef obj As cLinkField) As Range
    ' helper function
    On Error GoTo errRange
    Set getRangeKey = thisWb.Worksheets(obj.keyWorksheetName).Range(obj.keyRangeName)
    Exit Function
errRange:
    MsgBox ("[cUpdater:getRangeKey] Error in getting range. Ws: " & obj.keyWorksheetName & " rg: " & obj.keyRangeName)
    
End Function

Public Function getValue(ByRef obj As cLinkField)
    ' helper function
    getValue = thisWb.Worksheets(obj.WorksheetName).Range(obj.RangeName).value
End Function

Public Function getRangeRowCount(ByRef obj As cLinkField)
    ' helper function
    getRangeRowCount = thisWb.Worksheets(obj.WorksheetName).Range(obj.keyRangeName).Rows.count
End Function

Public Function getKey(ByRef obj As cLinkField)
    ' helper function
    getKey = thisWb.Worksheets(obj.keyWorksheetName).Range(obj.keyRangeName).value
End Function


' updates one SINGLE value to the DB
Public Sub updateDataToDB(ByVal dataType As String, ByVal keyType As String, ByVal tableName As String, _
               ByVal dataColumnName As String, ByVal dataValue, ByVal keyColumnName As String, ByVal keyValue)
    Dim matchedType As Boolean
    matchedType = False


        If keyType = "STR" And dataType = "STR" Then
            ' update STRING value to DB
            Call updateStringValueStringKey(tableName, dataColumnName, dataValue, keyColumnName, keyValue)
            matchedType = True
        End If
        
        
        If keyType = "STR" And dataType = "DOUBLE" Then
            ' update DOUBLE value to DB
            Call updateNumValueStringKey(tableName, dataColumnName, dataValue, keyColumnName, keyValue)
            matchedType = True
        End If
        
        If keyType = "STR" And dataType = "DATE" Then
            ' update DATE value to DB
            Call updateDateValueStringKey(tableName, dataColumnName, dataValue, keyColumnName, keyValue)
            matchedType = True
        End If
        
        If keyType = "STR" And dataType = "BOOL" Then
            ' update BOOLEAN value to DB
            Call updateNumValueStringKey(tableName, dataColumnName, convertStringToInt(dataValue), keyColumnName, keyValue)
            matchedType = True
        End If
        
        If Not matchedType Then
            MsgBox "updateFieldToDB" & vbCrLf & "Error:" & vbCrLf & "Type <key:" & keyType & _
            ",val:" & dataType & "> not supported? Value not updated", vbCritical, "Error"
        End If
        

End Sub


'updates one SINGLE value to the DB
Public Sub updateCellPairToDB(ByRef obj As cLinkField, ByRef rgKey As Range, ByRef rgValue As Range)
    If (rgKey Is Nothing) Or (rgValue Is Nothing) Then Exit Sub
    
    If rgKey.value = Empty Then
        Exit Sub
    End If
    
    Call updateDataToDB(obj.type_, obj.keyType_, obj.tableName, obj.columnName, rgValue.value, obj.keyColumnName, rgKey.value)
    
End Sub

'updates one SINGLE value to the DB
' SINGLE field - that is, one CELL
Public Sub updateFieldToDB(ByRef obj As cLinkField)
    
    Dim matchedType As Boolean
    matchedType = False

        If Not checkRangeExists(thisWb, obj.WorksheetName, obj.RangeName) Then Exit Sub
        If Not checkRangeExists(thisWb, obj.keyWorksheetName, obj.keyRangeName) Then Exit Sub
        
'        Call updateDataToDB(obj.type_, obj.keyType_, obj.tableName, obj.ColumnName, getValue(obj), obj.keyColumnName, getKey(obj))
         Call updateCellPairToDB(obj, getRangeKey(obj), getRangeData(obj))
         
End Sub





' check that all RANGEs EXIST
' optionally displays warnings
Public Function checkRanges(ByRef obj As cLinkList, Optional ByVal dispErrors As Boolean = True) As Boolean
' NOTE
' questa va bene sia per SINGLE-link (CELL) che per MULTI-link (COLUMN)
    Dim a As cLinkField
    Dim noError As Boolean
    noError = True
    
    For Each a In obj.oList
        
        'check VALUE RANGE
        If mod_Checks.checkRangeExists(thisWb, a.WorksheetName, a.RangeName, mod_Checks.DISP_NO_ERRORS) _
                                                                <> mod_Checks.RANGE_EXISTS Then
            If dispErrors Then
                MsgBox _
                    "Range not found: " & vbCrLf & _
                    "Worksheet: " & a.WorksheetName & vbCrLf & _
                    "Range: " & a.RangeName & vbCrLf & _
                    "on LinkID: " & a.linkID
            End If
            noError = False
        End If
    
        'check KEY RANGE
        If mod_Checks.checkRangeExists(thisWb, a.keyWorksheetName, a.keyRangeName, mod_Checks.DISP_NO_ERRORS) _
                                                                <> mod_Checks.RANGE_EXISTS Then
            If dispErrors Then
                MsgBox _
                    "Range not found: " & vbCrLf & _
                    "Worksheet: " & a.keyWorksheetName & vbCrLf & _
                    "Range: " & a.keyRangeName & vbCrLf & _
                    "on LinkID: " & a.linkID
            End If
            noError = False
        End If
    
    Next
    
    If dispErrors Then
        If Not noError Then
            MsgBox "Check of Ranges/Links completed with errors", vbCritical, "Check Failed"
        End If
    End If

    ' returns true or false
    checkRanges = noError
End Function



Public Function checkAllRangeValues(ByRef obj As cLinkList, Optional ByVal dispErrors As Boolean = True) As Boolean
    Dim a As cLinkField
    Dim noError, rangeTypeOk As Boolean
    
    ' cycle over links
    For Each a In obj.oList
        rangeTypeOk = checkMultiRangeType(getRange(a), a.type_, True, False)
    Next
    
End Function


' updates the data to the DB
' cycles over the Link list
Public Function updateXlsToDB(ByRef obj As cLinkList, Optional ByVal dispErrors As Boolean = True) As Boolean
    
    Dim a As cLinkField
    Dim noError, rangeTypeOk, dataTypeOk, keyTypeOk As Boolean
    noError = True
    Dim i, n As Long
    
    For Each a In obj.oList
        
        ' if it is SINGLE-CELL
        If a.linkType = "CELL" Then
            ' qui: check single o check multi..
            dataTypeOk = checkSingleRangeType(getRangeData(a), a.type_, True, False)
            keyTypeOk = checkSingleRangeType(getRangeKey(a), a.keyType_, True, False)
            'if type ok
            If dataTypeOk And keyTypeOk Then
                Call updateFieldToDB(a)
            End If
        End If
        
        ' if it is COLUMN
        If a.linkType = "COLUMN" Then
            n = getRangeRowCount(a)
            For i = 1 To n
                ' cycle.. single cell only
                dataTypeOk = checkSingleRangeType(getRangeData(a).Cells(i), a.type_, True, False)
                keyTypeOk = checkSingleRangeType(getRangeKey(a).Cells(i), a.keyType_, True, False)
                'if type ok
                If dataTypeOk And keyTypeOk Then
                    Call updateCellPairToDB(a, getRangeKey(a).Cells(i), getRangeData(a).Cells(i))
                End If
            Next
        End If
        
        ' if it is N_TO_1
        If a.linkType = "COL_N_TO_1" Then
            n = getRangeRowCount(a)
            
            ' data is SINGLE-cell
            dataTypeOk = checkSingleRangeType(getRangeData(a), a.type_, True, False)
            
            For i = 1 To n
                ' cycle.. key is MULTI-cell (COLUMN)
                keyTypeOk = checkSingleRangeType(getRangeKey(a).Cells(i), a.keyType_, True, False)
                'if type ok
                If dataTypeOk And keyTypeOk Then
                    Call updateCellPairToDB(a, getRangeKey(a).Cells(i), getRangeData(a))
                End If
            Next
        End If
        
        
        If (a.linkType <> "CELL") And (a.linkType <> "COLUMN") And (a.linkType <> "COL_N_TO_1") Then
            MsgBox "Warning! Unsupported linkType <" & a.linkType & ">" & vbCrLf & vbCrLf & "cUpdater:updateXlsToDB"
        End If
    
    Next
    
    
    updateXlsToDB = noError
    
End Function





