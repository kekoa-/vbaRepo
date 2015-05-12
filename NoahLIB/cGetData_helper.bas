VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGetData_helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public linkList As Collection

Option Compare Text

'aa
Public Sub init()
    Dim rs As Recordset
    Dim oElem As cGetData_element
    
    On Error Resume Next
    
    Set linkList = New Collection
    Call DB_Utilities.execTableSQL_withRS("select strCode,strCommand,strTableName, strColumnName, strKeyName, " & _
                                          "strKeyType from tblgetdata_links where not(strkeytype is null);", rs)
    rs.MoveFirst
    
    While Not rs.EOF
        Set oElem = New cGetData_element
        oElem.code = rs("strCode").value
        oElem.command = rs("strCommand").value
        oElem.tableName = rs("strTableName").value
        oElem.columnName = rs("strColumnName").value
        oElem.keyName = rs("strKeyName").value
        oElem.keyType = rs("strKeyType").value
        
        Call linkList.add(oElem)
    
        rs.MoveNext
    Wend
    
    On Error GoTo 0

    
End Sub

Public Function getData(ByVal code As String, _
    ByVal command As String, _
    ByVal key As String)

    Dim o As cGetData_element
    

    For Each o In linkList
        If (o.code = code) And (o.command = command) Then
            If (o.keyType = "STR") Then
                getData = modDBInterface.getScalarStringKey(o.tableName, o.columnName, o.keyName, key)
                Exit Function
            End If
            If (o.keyType = "INT") Then
                getData = modDBInterface.getScalarNumKey(o.tableName, o.columnName, o.keyName, key)
                Exit Function
            End If
            
        End If
    Next
    
    getData = "## Invalid Input"

End Function


