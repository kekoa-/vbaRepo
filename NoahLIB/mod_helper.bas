Attribute VB_Name = "mod_helper"
Option Explicit

Public Const WS_ERROR = -1
Public Const WS_CATBOND = 1
Public Const WS_CATSWAP = 2
'


' called within a SwapAnalysis or CatBond Analysis workbook, returns the Scheda-type
Public Function checkSchedaType(ByRef wb As Workbook) As Long
Dim rg As Range

On Error GoTo checkNext
    ' check if is CatBond
    Set rg = wb.Worksheets("Summary").Range("rng_strAsset_CUSIP")
    checkSchedaType = WS_CATBOND
    On Error GoTo 0
    Exit Function


checkNext:
On Error GoTo checkNext2

    Set rg = wb.Worksheets("Summary").Range("rng_Layer_Name")
    checkSchedaType = WS_CATSWAP
    On Error GoTo 0
    Exit Function

checkNext2:
    checkSchedaType = WS_ERROR
    On Error GoTo 0
    Exit Function

End Function

' gets the UMR of catSwap Program or Code fo CatBond
Public Function getOwnerCode(ByRef wb As Workbook) As String

    getOwnerCode = ""
    If checkSchedaType(wb) = WS_CATBOND Then
        getOwnerCode = wb.Worksheets("Summary").Range("rng_strAsset_Code")
    End If
    If checkSchedaType(wb) = WS_CATSWAP Then
        getOwnerCode = wb.Worksheets("Summary").Range("rng_UMR")
    End If

End Function

' gets owner type string
Public Function getOwnerType(ByRef wb As Workbook) As String

    getOwnerType = ""
    If checkSchedaType(wb) = WS_CATBOND Then
        getOwnerType = "CB"
    End If
    If checkSchedaType(wb) = WS_CATSWAP Then
        getOwnerType = "RE"
    End If

End Function

' writes the column names of the recordset
Public Sub writeRecordsetColumnHeader(ByRef rg As Range, ByRef rs As Recordset)
    If (rs Is Nothing) Or (rg Is Nothing) Then Exit Sub
    
    Dim i, n As Long
    n = rs.Fields.count
    For i = 0 To n - 1
        rg.Cells(1, i + 1).value = rs.Fields.Item(i).Name
    Next


End Sub


Public Sub safeCopyFromRecordset(ByRef rg As Range, ByRef rs As Recordset)
    If (rs Is Nothing) Or (rg Is Nothing) Then Exit Sub
    Call rg.CopyFromRecordset(rs)
End Sub


' format range rg based on the columns names
Public Sub formatRangesByName(ByRef rg As Range, ByRef rg_header As Range)
    If (rg Is Nothing) Or (rg_header Is Nothing) Then Exit Sub
    
    Dim i, n1, n2, n As Long
    Dim strHeader As String
    
    n1 = rg.Columns.count
    n2 = rg_header.Columns.count
    If n1 > n2 Then
        n = n1
    Else
        n = n2
    End If
    
    
    For i = 1 To n
        strHeader = rg_header.Cells(1, i).value
        If Left(strHeader, 3) = "int" Then
            rg.Columns(i).NumberFormat = "General"
        ElseIf Left(strHeader, 3) = "boo" Then
            rg.Columns(i).NumberFormat = "General"
        ElseIf Left(strHeader, 3) = "dat" Then
            rg.Columns(i).NumberFormat = "yyyy-mm-dd"
        ElseIf Left(strHeader, 3) = "str" Then
            rg.Columns(i).WrapText = False
        ElseIf strHeader = "dblSpreadNROL" Then
            rg.Columns(i).NumberFormat = "0.000%"
        ElseIf strHeader = "dblELBookMulti" Then
            rg.Columns(i).NumberFormat = "0.000%"
        ElseIf Left(strHeader, 3) = "dbl" Then
            rg.Columns(i).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        End If
        
    Next
End Sub



Public Function get_AssetCode_byNick(ByVal assetNick As String)
    get_AssetCode_byNick = DB_Utilities.execScalarSQL("SELECT strCode FROM tblasset WHERE strNick='" & assetNick & "';")
End Function


' check if the string contains special characters, that is, INVALID characters.
Public Function hasSpecialCharacters(ByVal s As String)
    Dim i As Integer, n As Integer, val As Integer
    Dim c
    
    hasSpecialCharacters = False
    n = Len(s)
    For i = 1 To n
        c = Mid(s, i, 1)
        val = AscW(c)
        If (val > 127) Or (c = "'") Or (c = " ") Then
            hasSpecialCharacters = True
        End If
'        MsgBox val
    Next
    
End Function

' checks that this range contains no special characters, otherwise gives an error, highlights the range, and exits
Public Sub check_ASCIIcodeForKey_intoRange_orExit(ByRef rg As Range)
    If rg Is Nothing Then Exit Sub
    Dim res
    res = hasSpecialCharacters(rg.value)
    If (res = True) Then
        rg.Parent.Activate
        rg.Select
        rg.Interior.Color = RGB(255, 0, 0)
        MsgBox "This range contains invalid characters, please check the value. Exit."
        End
    End If
End Sub


Public Sub testSpecialCharacters()
Dim res
Call check_ASCIIcodeForKey_intoRange_orExit(Range("rng_Nick"))
'    res = hasSpecialCharacters()
'    res = hasSpecialCharacters("123")
'    res = hasSpecialCharacters("asd12")
    

End Sub




Public Sub cleanTarget_down(ByRef rg_target As Range, ByVal col_target_count As Integer)
    Dim ws As Worksheet
    Dim rg_clean As Range
    
    If Not rg_target Is Nothing Then
        If rg_target.Cells(1, 1).value = "" Then Exit Sub
        Set ws = rg_target.Parent
        Set rg_clean = ws.Cells.Range(rg_target.Cells(1, 1), rg_target.Cells(1, 1).End(xlDown).Cells(1, col_target_count))
        rg_clean.ClearContents
        Set rg_clean = Nothing
        Set ws = Nothing
    End If

End Sub








