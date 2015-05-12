VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLinkLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

' carica la lista di campi Excel<-->DB filtrando per strWsType
' legge la lista da DB
Public Sub loadLinkList(ByRef obj As cLinkList, ByVal strWsType As String)

    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Dim newLF As cLinkField
    
    obj.init
    
    ' get field list from DB
    strSQL = "select * from tbllinkfields where strWsType='" & strWsType & "'   ;"
    'execute query
    Call DB_Utilities.execTableSQL_withRS(strSQL, rs)
    
    'cycle over Recordset
    While Not rs.EOF
        Set newLF = New cLinkField
        On Error GoTo errSkipRow
        ' get attributes from the record
            newLF.linkID = rs("intID")
            newLF.tableName = rs("strTableName")
            newLF.keyColumnName = rs("strKeyColumnName")
            newLF.keyType_ = rs("strKeyType")
            newLF.keyWorksheetName = rs("strKeyWsName")
            newLF.keyRangeName = rs("strKeyRangeName")
            newLF.columnName = rs("strColumnName")
            newLF.type_ = rs("strType")
            newLF.WorksheetName = rs("strWsName")
            newLF.RangeName = rs("strRangeName")
    '  TODO: RANGE_TYPE è CELL o RANGE
            newLF.linkType = rs("strLinkType")
            
            obj.oList.add newLF
            
            GoTo resumeCycle

errSkipRow:
'    On Error GoTo 0
    MsgBox "Error in <cLinkLoader:loadCBFieldsList>, exit now"
    End
resumeCycle:
        
        rs.MoveNext
    Wend
    
On Error GoTo 0

End Sub


' loads LinkList for CatBond XLS
Public Sub loadCBFieldsList(ByRef obj As cLinkList)
    Call loadLinkList(obj, "SCHEDA_CB")
End Sub

' loads LinkList for CatSwap XLS
Public Sub load_CatSwap_FieldsList(ByRef obj As cLinkList)
    Call loadLinkList(obj, "SCHEDA_CATSWAP")
End Sub



