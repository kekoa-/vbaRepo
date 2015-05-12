VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRMS_helper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private analysisIdList() As String
Private regionList() As String
Private perilList() As String


Public analysisCount As Integer
Public intRMS As Integer
Public perspcode As String

' to be called before using this object
Public Sub init()
    analysisCount = 0
    ReDim analysisIdList(100)
    ReDim regionList(100)
    ReDim perilList(100)
End Sub

' adds the analysis to the lists
Public Sub addAnalysis(ByVal analysisId As Integer, ByVal region As String, ByVal peril As String)
    analysisCount = analysisCount + 1
    analysisIdList(analysisCount) = analysisId
    regionList(analysisCount) = region
    perilList(analysisCount) = peril
End Sub

' sets the perspcode
Public Sub setPerspcode(ByVal perspCode_in As String)
    perspcode = perspCode_in
End Sub


' cnn is the open connection to the SQL Server
' this sub packs the selected analyses into a temp table
Public Sub generateCombinedAnalysis(ByRef cnn As ADODB.Connection)
    Dim i As Integer
    Dim s1 As String, strSQL As String
    
    If intRMS = 0 Then
        ' check that the destination RMS analysis ID is set
        Call MsgBox("[cRMS_helper:generateCombinedAnalysis] invalid value, intRMS is ZERO", vbExclamation)
        Exit Sub
    End If
    If perspcode = "" Then
        ' check that perspcode is set
        Call MsgBox("[cRMS_helper:generateCombinedAnalysis] invalid value, perspcode is empty", vbExclamation)
        Exit Sub
    End If
    If analysisCount = 0 Then
        ' check that at least one analysis has been selected
        Call MsgBox("[cRMS_helper:generateCombinedAnalysis] invalid value, analysisCount is ZERO, no analyses selected?", vbExclamation)
        Exit Sub
    End If
    
    ' empties temp tables
    On Error Resume Next
    s1 = " DROP TABLE ##ttbl_RMSExport;"
    cnn.Execute (s1)
    s1 = " DROP TABLE ##ttbl_RMSExport2;"
    cnn.Execute (s1)
    On Error GoTo 0
    
    ' for each analysis
    For i = 1 To analysisCount
        If i = 1 Then   ' creates the temp table and gets the first data
            s1 = sqlRMS_QueryELT_forHelper_createTemp(intRMS, regionList(i), perilList(i), analysisIdList(i), perspcode)
            cnn.Execute (s1)
        Else            ' append data
            s1 = sqlRMS_QueryELT_forHelper_insertTemp(intRMS, regionList(i), perilList(i), analysisIdList(i), perspcode)
            cnn.Execute (s1)
        End If
    Next
    
    ' finally, sum by event
    s1 = sqlRMS_QueryELT_forHelper_packTemp()
    cnn.Execute (s1)
End Sub


' this sub retrieves the data from the temp table
Public Sub retrieveCombinedData(ByRef cnn As ADODB.Connection, ByRef rs As ADODB.Recordset)
    ' gets the data from the temp table to the recordset
    Call execTableSQL_withConn2(" select * from ##ttbl_RMSExport2;", cnn, rs)
End Sub















