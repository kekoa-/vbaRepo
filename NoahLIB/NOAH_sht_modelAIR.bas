Attribute VB_Name = "NOAH_sht_modelAIR"
Option Explicit

'##############################################################################################
' ROUTINES, FORMS AND FUNCTIONS RELATIVE TO THE AIR SHEET
'   [1]  Sub_AIR_LoadDBNames        (to be checked)
'   [2]  Sub_AIR_RetrieveData       (to be developed)
'   [3]  Sub_AIR_ImportELT          (to be reviewed)
'   [4]  Sub_AIR_ShowELT            (to be developed)
'   [5]  Sub_AIR_DeleteELT          (to be developed)
'   [6]  Sub_AIR_CalculateRisk      (to be developed)
'   [7]  Sub_AIR_SubmitOEP          (to be developed)
'   [8]  Form_AIR_ImportELT         (to be reviewed)
'   [9]  Form_AIR_ShowELT           (to be developed)
'   [10] Form_AIR_DeleteELT         (to be developed)
'##############################################################################################


Public Sub Sub_AIR_LoadDBNames(ByRef myWb As Workbook, _
                               ByRef mySht As Worksheet, _
                               ByRef myCmb As ComboBox, _
                               Optional ByVal BondOnly As Boolean = False, _
                               Optional ByVal SwapOnly As Boolean = False)
'==============================================================================================
' [1] SUB: TO POPULATE COMBOBOX WITH DB NAMES
'==============================================================================================
' Input:   #myWb    :  the reference to the workbook which calles this routine
'          #myForm  :  the reference to the form which called this routine
' Output:  -
' Descr:   in the relevant sheet populates the ComboBox containing the names of those
'          Companies which have been created within the CaTrader Database
'
' Version:
' IL - 17/07/2014
'----------------------------------------------------------------------------------------------

' Set Global Variables
Set wb1 = myWb
Set sh1 = mySht

' Defines variable
Dim SQLtype As String
Dim SQLname As String
Dim DBname As String
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim bb As ComboBox
        
        ' Inizializes values
        SQLname = mySht.Range("rng_AIR_SQLserver")
        DBname = "AirCT2Exp"

        ' Open the connection to the Database
        SQLtype = "AIR"
        Set cnn = New ADODB.Connection
        cnn.ConnectionString = strConn(SQLtype, SQLname, DBname)
        cnn.Open

        ' Calls the function to create the proper SQL query
        strSQL = sqlAIR_AvailCompany(BondOnly:=BondOnly, _
                                     SwapOnly:=SwapOnly)
        ' Retrive the available analyses
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cnn, adOpenDynamic, adLockOptimistic
        
        ' Clears and Populates the ComboBox
        myCmb.value = ""
        With myCmb
            .Clear
            rs.MoveFirst
            Do While rs.EOF = False
                myCmb.AddItem rs.Fields("strName").value
                rs.MoveNext
            Loop
        End With
        
End Sub
'==============================================================================================



Public Sub Sub_AIR_RetrieveData(ByRef myWb As Workbook, _
                                ByRef mySht As Worksheet)
'==============================================================================================
' [2] SUB: TO RETRIEVE STORED ANALYSES
'==============================================================================================
' Input:   #myWb    :  the reference to the workbook which calles this routine
'          #myForm  :  the reference to the form which called this routine
' Output:  -
' Descr:   retrieves the list of analyses that have been loaded for this CatBond/CatSwap program
' Version:
' DB 2015-02-02
'==============================================================================================
 
Dim numAnalyses As Long
Dim ownerCode, ownerType As String

Dim rg1, rg2, rg3, rg4 As Range
Dim rs As Recordset

On Error GoTo errFunction

If checkSchedaType(myWb) = WS_ERROR Then
    MsgBox "Error, unable to find SCHEDA_TYPE, exit", vbCritical
    End
End If

    ' get analysis owner data
    ownerCode = getOwnerCode(myWb)
    ownerType = getOwnerType(myWb)

numAnalyses = modDBInterface_AIRData.getAnalysesCount_byOwner(ownerCode)
'MsgBox "" & numAnalyses & " analyses found for <" & ownerCode & "," & ownerType & " >"


' add rows if needed
On Error GoTo errFunction
Set rg1 = Range("rng_AIR_AnalysesID")
Set rg2 = Range("rng_AIR_AnalysesDescr")
Set rg3 = rg2.Cells(rg2.Rows.count - 1, 1)
Set rg4 = mySht.Cells.Range(rg1.Cells(1, 1), rg3)
rg4.Select

'resize(adds) range and clears it
Call mod_rng_resize.range_insertRows_worksheet(rg4, numAnalyses)
rg1.ClearContents
rg2.ClearContents
rg1.Cells(1, 1).Select

If modDBInterface_AIRData.getAnalysisList_byOwner(rs, ownerCode) = True Then
    ' get analyses list
    rg4.CopyFromRecordset rs
End If


On Error GoTo 0
Exit Sub
errFunction:
On Error GoTo 0
MsgBox "[NOAH_sht_modelAIR:Sub_AIR_RetrieveData] Error: " & Err.Description

End Sub
'==============================================================================================



Public Sub Sub_AIR_ImportELT(ByRef myWb As Workbook, _
                             ByRef mySht As Worksheet, _
                             ByRef myForm As Object)
'==============================================================================================
' [3] SUB: TO IMPORT ELT
'==============================================================================================
' Input:   #myWb    :  the reference to the workbook which calles this routine
'          #myForm  :  the reference to the form which called this routine
' Output:  -
' Descr :  Handles the "SUBMIT" button_click event in the "Import ELT" UserForm.
'          Retrieves the data, writes CSVs, and imports data to MySQL.
' Version:
' DB 2015-01-30
'----------------------------------------------------------------------------------------------
    Dim idELT As Long, i As Long
    Dim guidIdCond As String
    Dim sql, ownerCode As String, conditionName As String
    
    ' get owner code
    ownerCode = getOwnerCode(myWb)
    ' Import into the MySQL database the selected analyses
    With myForm.tbl_Analyses
        i = 1
        ' CYCLE over the table in the "Import ELT" UserForm
        Do While .Columns.Cells(i, 2) <> ""
            'IF analysis is selected for import:
            If .Columns.Cells(i, 1).value <> "" Then
                ' get condition Name
                conditionName = .Columns.Cells(i, 3).value
                ' gets GUID
                guidIdCond = .Columns.Cells(i, 4).value
                
                ' inserts an empty Catrader ELT
                Call insertCatraderCondition(conditionName, idELT)
                ' set ELT owner
                Call modDBInterface.updateStringValueNumKey("tblCondition", "strOwner", ownerCode, "intID", idELT)
                ' updates GUID to th DB
                Call modDBInterface.updateStringValueNumKey("tblcondition", "strguidcondition", guidIdCond, "intid", idELT)
                ' updates Condition data
                Call updateConditionData(guidIdCond)
                
                ' gets the ELT points into the DB
                Call import_ELT_Catrader(idELT, guidIdCond)
                
            End If
            i = i + 1
        Loop
    End With
    
    ' refresh data in the worksheet
    Call Sub_AIR_RetrieveData(myWb, mySht)
    
    MsgBox "Importing done."
End Sub
'==============================================================================================



Public Sub Sub_AIR_ShowELT(ByRef myWb As Workbook, _
                           ByRef mySht As Worksheet)
'==============================================================================================
' [4] SUB: TO SHOW STORED ELT
'==============================================================================================

    MsgBox "Routine 'Sub_AIR_ShowELT' is Under Construction"
    
End Sub
'==============================================================================================



Public Sub Sub_AIR_DeleteELT(ByRef myWb As Workbook, _
                              ByRef mySht As Worksheet)
'==============================================================================================
' [5] SUB: TO DELETE STORED DATA
'==============================================================================================

    MsgBox "Routine 'Sub_AIR_DeleteELT' is Under Construction"

End Sub
'==============================================================================================


' "Estimate Risk" button in the AIR Worksheet
Public Sub Sub_AIR_CalculateRisk(ByRef myWb As Workbook, _
                                 ByRef mySht As Worksheet)
'==============================================================================================
' [6] SUB: TO CALCULATE RISK MEASURES
'==============================================================================================
' Sets global variables
Set wb1 = myWb
Set sh1 = mySht

Dim s, strSQL, res, schedaType
Dim rg As Range
Dim i, n, analysisId, assetNick, occEL, aggEL, attProb, exhProb, attProbAgg, exhProbAgg, ELccy, StddevCcy

Call getAssetList_AIR(myWb:=wb1, _
                      mySht:=sh1)

'MsgBox "Refresh data.."
s = mod_helper.getOwnerCode(myWb)
schedaType = mod_helper.getOwnerType(myWb)

Set rg = mySht.Range("rng_AIR_LayerName")
rg.ClearContents
strSQL = "SELECT strnick, intCondition from tblasset where strowner='" & s & "' order by intassetnum;"
res = mod_btn_GetData.putRsIntoRange_addByRow(strSQL, rg)


' for each row
For i = 1 To rg.Rows.count
    ' get asset nick
    assetNick = rg.Cells(i, 1).value
    ' get the ID of associated analysis
    analysisId = rg.Cells(i, 2).value
    
    If (assetNick <> "") And (analysisId <> "") Then
  
        DB_Utilities.execCommandSQL ("call prcGetAggregateELforCondition(" & analysisId & ",@res1);")
        aggEL = modDBInterface.getVar("@res1")
        DB_Utilities.execCommandSQL ("call prcGetOccurrenceELforCondition(" & analysisId & ",@res2);")
        occEL = modDBInterface.getVar("@res2")
        DB_Utilities.execCommandSQL ("call prcGetOccurrenceAttachmentProbability(" & analysisId & ",@res3);")
        attProb = modDBInterface.getVar("@res3")
        DB_Utilities.execCommandSQL ("call prcGetOccurrenceExhaustionProbability(" & analysisId & ",@res4);")
        exhProb = modDBInterface.getVar("@res4")
        DB_Utilities.execCommandSQL ("call prcGetAggregateELCcyforCondition(" & analysisId & ",@res5);")
        ELccy = modDBInterface.getVar("@res5")
        DB_Utilities.execCommandSQL ("call prcGetAggregateStddevCcyforCondition(" & analysisId & ",@res6);")
        StddevCcy = modDBInterface.getVar("@res6")
        
        DB_Utilities.execCommandSQL ("call prcGetAggregateAttProbforCondition(" & analysisId & ",@res7);")
        attProbAgg = modDBInterface.getVar("@res7")
        DB_Utilities.execCommandSQL ("call prcGetAggregateExhProbforCondition(" & analysisId & ",@res8);")
        exhProbAgg = modDBInterface.getVar("@res8")
      
        rg.Cells(i, 6).value = attProb
        rg.Cells(i, 7).value = occEL
        rg.Cells(i, 8).value = exhProb
        
        rg.Cells(i, 9).value = attProbAgg
        rg.Cells(i, 10).value = aggEL
        rg.Cells(i, 11).value = exhProbAgg
        
'        If (schedaType = "RE") Then
        rg.Cells(i, 12).value = ELccy
        rg.Cells(i, 13).value = StddevCcy
            
'        End If
        
    End If
Next


End Sub
'==============================================================================================

Public Sub Sub_AIR_SubmitOEP(ByRef myWb As Workbook, ByRef mySht As Worksheet)
    Call Sub_AIR_ShowUtilitiesForm(myWb, mySht)
End Sub

Public Sub Sub_AIR_ShowUtilitiesForm(ByRef myWb As Workbook, _
                                     ByRef mySht As Worksheet)
'==============================================================================================
' [7] SUB: TO SUBMIT OEP
'==============================================================================================

' Sets global variables
Set wb1 = myWb
Set sh1 = mySht
userForm_AIR_Analysis.Show

End Sub
'==============================================================================================



Public Sub Form_AIR_ImportELT(ByRef myWb As Workbook, _
                              ByRef mySht As Worksheet)
'==============================================================================================
' [8] FORM: TO IMPORT ELT
'==============================================================================================
' Input:   #myWb    :  the reference to the workbook which calles this routine
'          #myForm  :  the reference to the form which called this routine
' Descr:   initialized the proper input required by the Import_ELT and runs it
'----------------------------------------------------------------------------------------------

' Sets global variables
Set wb1 = myWb
Set sh1 = mySht

' Defines variables
Dim myForm As Object
Dim SQLtype As String
Dim SQLname As String
Dim DBname As String
Dim CompanyName As String
Dim DBlist As String
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim bb As OptionButton
Dim i As Integer
    
        ' Inizializes values
        Set myForm = userForm_AIR_ImportELT
        SQLname = mySht.Range("rng_AIR_SQLserver")
        CompanyName = mySht.OLEObjects("cmb_CompanyList").Object.value
        DBname = "AirCT2Exp"

        ' Open the connection to the Database
        SQLtype = "AIR"
        Set cnn = New ADODB.Connection
        cnn.ConnectionString = strConn(SQLtype, SQLname, DBname)
        cnn.Open

        ' Retrive the available analyses
        strSQL = sqlAIR_AvailContract(CompanyName)
        Set rs = New ADODB.Recordset
        'Cells(3, 2) = strSQL
        rs.Open strSQL, cnn, adOpenDynamic, adLockOptimistic

        ' Load the form to import ELT
        Load myForm
        myForm.txt_SQLname = SQLname
        myForm.txt_CompanyName = CompanyName
        
        With myForm.tbl_Analyses
            i = 1
            .Cells.ClearContents
            Do While rs.EOF = False
                .Cells(i, 2) = rs.Fields("contractName").value
                .Cells(i, 3) = rs.Fields("conditionName").value
                .Cells(i, 4) = rs.Fields("conditionID").value
                i = i + 1
                rs.MoveNext
            Loop
            .Cells(1, 1).Select
        End With
        
        ' Shows the form
        myForm.Show

End Sub
'==============================================================================================







Public Sub Form_AIR_ShowELT(ByRef myWb As Workbook, _
                            ByRef mySht As Worksheet)

Dim rg, rg2, rg3 As Range
Dim intID As Long
Dim newWs As Worksheet
Dim rs As Recordset

'On Error GoTo errFunction
Set rg = Range("rng_AIR_AnalysesID")

Set rg2 = ActiveCell.Cells(1, 1)

Set rg3 = mySht.Cells(rg2.Row, rg.column)
rg3.Select

If Not mod_Checks.checkSingleRangeType(rg3, "INT") Or (rg3.value = "") Then
    MsgBox "Invalid ELT Id found <" & rg3.value & "> , exit."
    End
End If

MsgBox "Showing the ELT with ID <" & rg3.value & "> ", vbInformation
intID = rg3.value
'Shell ("notepad.exe " & "K:\Shared Modeling\ANALISI FUNZIONALI\NOAH\Database\CSV_repo\ELT_Catrader_" & intID & ".csv")
Set newWs = myWb.Worksheets.add()
newWs.Activate
Call DB_Utilities.execTableSQL_withRS("select * from tbleltcatrader where intcondition=" & intID, rs)
Call writeRecordsetColumnHeader(newWs.Cells(1, 2), rs)
Call newWs.Cells(2, 2).CopyFromRecordset(rs)


On Error GoTo 0
Exit Sub
errFunction:
On Error GoTo 0
MsgBox "[NOAH_sht_modelAIR:Form_AIR_ShowELT] Error: " & Err.Description


End Sub












Public Sub Form_AIR_ShowELT2(ByRef myWb As Workbook, _
                            ByRef mySht As Worksheet)
'==============================================================================================
' [9] FORM: TO SELECT THE ELT TO BE IMPORTED
'==============================================================================================
' Input:    #myWb   :  the workbook which is calling the routine
'           #mySht  :  the worksheet which is calling the routine
' Output:   -
' Descr:    Shows the form to select the ELTs (within the Katarsis MySQL database) to be
'           imported into this workbook
' Vers:
' 24.07.2014 - IL
'----------------------------------------------------------------------------------------------

Dim rngID As Range, rngDescr As Range
Dim count, i As Long
Dim myForm As Object

    MsgBox "Routine 'Form_AIR_ShowELT' is Under Construction"
    Exit Sub

    ' Initialize variable
    Set myForm = userForm_ShowELT
    Set rngID = mySht.Range("rng_AIR_AnalysesID")
    Set rngDescr = mySht.Range("rng_AIR_AnalysesDescr")

    ' Load the form to select the ELT
    Load myForm
    With myForm.list_SelectELT
        .ColumnCount = 4
        .ColumnWidths = .Width * 0.01 & ";" & .Width * 0.1 & ";" & .Width * 0.1 & ";" & .Width * 0.55
    End With
    
    ' Populates the ListBox showing the available ELT
    count = 0
    For i = 1 To rngID.Rows.count
        If rngID.Cells(i, 1) <> "" Then
            With myForm.list_SelectELT
                .AddItem
                .list(i - 1, 1) = rngID.Cells(i, 1)
                .list(i - 1, 2) = rngPort.Cells(i, 1)
                .list(i - 1, 3) = rngDescr.Cells(i, 1)
            End With
            count = count + 1
        End If
    Next
    
    ' In case there is no available ELT shows an Error Message
    If count <> 0 Then
        myForm.Show
    Else
        Unload myForm
        MsgBox Prompt:="Attention! The list of available analyses is empty.", _
               Title:="NOAH - Error"
    End If

End Sub
'==============================================================================================



Public Sub Form_AIR_DeleteELT(ByRef myWb As Workbook, _
                              ByRef mySht As Worksheet)
'==============================================================================================
' [10] FORM: TO DELETE STORED ELT
'==============================================================================================
' Input:    #myWb   :  the workbook which is calling the routine
'           #mySht  :  the worksheet which is calling the routine
' Output:   -
' Descr:    -
' Vers:
' -
'==============================================================================================

Dim rg, rg2, rg3 As Range
Dim intID As Long

'On Error GoTo errFunction
Set rg = Range("rng_AIR_AnalysesID")

Set rg2 = ActiveCell.Cells(1, 1)

Set rg3 = mySht.Cells(rg2.Row, rg.column)
rg3.Select

If Not mod_Checks.checkSingleRangeType(rg3, "INT") Or (rg3.value = "") Then
    MsgBox "Invalid ELT Id found <" & rg3.value & "> , exit."
    End
End If

If (MsgBox("Delete the ELT with ID <" & rg3.value & "> ?", vbYesNo) = vbYes) Then
    intID = rg3.value
    Call modDBInterface.deleteNumKey("tblCondition", "intId", intID)
    rg3.Cells(1, 1).ClearContents
    MsgBox "Deleted"
End If

' refresh the "Available ELT" table
Call Sub_AIR_RetrieveData(myWb, mySht)

On Error GoTo 0
Exit Sub
errFunction:
On Error GoTo 0
MsgBox "[NOAH_sht_modelAIR:Form_AIR_DeleteELT] Error: " & Err.Description


End Sub
'==============================================================================================





' note, imports ALL the ELTs
Public Sub Form_AIR_ImportELT_ALL(ByRef myWb As Workbook, _
                                  ByRef mySht As Worksheet)
    ' Sets global variables
    Set wb1 = myWb
    Set sh1 = mySht
    
    ' Defines variables
    Dim SQLtype As String
    Dim SQLname As String
    Dim DBname As String
    Dim CompanyName As String
    Dim DBlist As String
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String

    Dim idELT As Long, conditionName, guidIdCond, ownerCode, i As Integer
    
    
    If MsgBox("Import ALL the AIR analyses?", vbYesNo) <> vbYes Then
        MsgBox "Canceled, exit."
        End
    End If
    
    SQLname = mySht.Range("rng_AIR_SQLserver")
    CompanyName = mySht.OLEObjects("cmb_CompanyList").Object.value
    DBname = "AirCT2Exp"
    
    ' Open the connection to the Database
    SQLtype = "AIR"
    Set cnn = New ADODB.Connection
    cnn.ConnectionString = strConn(SQLtype, SQLname, DBname)
    cnn.Open
    
    ' Retrive the available analyses
    strSQL = sqlAIR_ALLContract()
    Set rs = New ADODB.Recordset
    'Cells(3, 2) = strSQL
    rs.Open strSQL, cnn, adOpenDynamic, adLockOptimistic
    rs.MoveFirst
    
    Cells.Range("k30").CopyFromRecordset rs
    
    ownerCode = getOwnerCode(myWb)
    rs.MoveFirst
    i = 1
    ' Import into the MySQL database the selected analyses
    Do While rs.EOF = False
        ' get condition Name
        conditionName = rs("conditionName").value
        ' gets GUID
        guidIdCond = rs("conditionID").value
        
        ' inserts an empty Catrader ELT
        Call insertCatraderCondition(conditionName, idELT)
        ' set ELT owner
        Call modDBInterface.updateStringValueNumKey("tblCondition", "strOwner", ownerCode, "intID", idELT)
        ' updates GUID to th DB
        Call modDBInterface.updateStringValueNumKey("tblcondition", "strguidcondition", guidIdCond, "intid", idELT)
        ' updates Condition data
        Call updateConditionData(guidIdCond)
        
        ' gets the ELT points into the DB
        Call import_ELT_Catrader(idELT, guidIdCond, False)
        
        Cells.Range("i30").value = "Loading Condition " & i
        
        i = i + 1
        rs.MoveNext
    Loop

    MsgBox "Done"

End Sub
'==============================================================================================





Public Sub getAssetList_AIR(ByRef myWb As Workbook, _
                            ByRef mySht As Worksheet)
'==============================================================================================
' SUB: TO GET THE LIST OF ASSET (retrieves asset list, and puts them into the appropriate range)
'==============================================================================================

Dim strSQL, res, rg As Range, rg1 As Range
Dim ownerCode, ownerType As String

    Set wb1 = myWb
    Set sh1 = mySht

    ' Checks if the current Sheet is one of those containing Risk Measures
    If checkSchedaType(wb1) = WS_ERROR Then
        ' If not prints and error and exits
        MsgBox "Error, unable to find SCHEDA_TYPE, exit", vbCritical
        End
    End If

    ' Detects the type of model to be used (according to the sheet name)
    ownerCode = getOwnerCode(wb1)
    ownerType = getOwnerType(wb1)
    
    ' Calls the routine to delete previous Risk Measures
    Call clearRiskMeasuresAir(myWb:=ActiveWorkbook, _
                              mySht:=ActiveSheet)
    
    Set rg = sh1.Range("rng_AIR_LayerName")
    Set rg1 = sh1.Range("rng_AIR_LayerStructure")
    
    ' Retrives data from the Datebase
    strSQL = "SELECT strnick, intCondition from tblasset where strowner='" & ownerCode & "' order by intassetnum;"
    res = mod_btn_GetData.putRsIntoRange_addByRow(strSQL, rg)
    
    ' If the resulting query is void prints an error
    If (res = 0) Then
        Call MsgBox("No assets are associated to this Worksheet/CatSwapProgram/CatBond." & vbCrLf & _
            "Swap Analysis/Security Analysis has not been submitted yet?" & vbCrLf & vbCrLf & _
            "Please click <Submit Data> in the <Summary> worksheet first.", vbInformation)
    End If
    
    ' Retrice Limit and Trigger values for a UNL-type asset
    If (ownerType = "RE") Then
        strSQL = "SELECT dblLimit, dblDeductible from tblcatswaplayer where strProgramUMR='" & ownerCode & "' order by intlayernum;"
        res = mod_btn_GetData.putRsIntoRange_addByRow(strSQL, rg1)
    End If
    
    ' Retrice Limit and Trigger values for a CatBond-type asset
    If (ownerType = "CB") Then
        strSQL = "SELECT dblLossLevelExhaustion-dblLossLevelTrigger, dblLossLevelTrigger from tblcatbondinfo where strAssetCode='" & ownerCode & "' ;"
        res = mod_btn_GetData.putRsIntoRange_addByRow(strSQL, rg1)
    End If

End Sub
'==============================================================================================



Public Sub clearRiskMeasuresAir(ByRef myWb As Workbook, _
                                ByRef mySht As Worksheet)
'==============================================================================================
' SUB: TO CLEAR RISK MEASURES FROM THE AIR SHEET
'==============================================================================================
    
    mySht.Range("rng_AIR_LayerName").ClearContents
    mySht.Range("rng_AIR_LayerGroup").ClearContents
    mySht.Range("rng_AIR_LayerStructure").ClearContents
    mySht.Range("rng_AIR_RiskMeasure").ClearContents
    mySht.Range("rng_AIR_RiskAmount").ClearContents
    mySht.Range("rng_AIR_LayerType").ClearContents
    On Error Resume Next
    mySht.Range("rng_AIR_RiskMeasureAggr").ClearContents
    On Error GoTo 0
    
End Sub
'==============================================================================================


Public Sub AIR_LinkAnalysisToAsset(ByRef myWb As Workbook, _
                                   ByRef mySht As Worksheet)
'==============================================================================================
' SUB: TO CLEAR RISK MEASURES FROM THE AIR SHEET
'==============================================================================================

Dim i, nickName
Dim rg_asset_nick, rg_analysis_id As Range
Dim analysisId As Long
Dim count As Integer

    If MsgBox("Associate the selected analyses to the assets?", vbQuestion + vbYesNo) <> vbYes Then
        MsgBox "Canceled, exit"
        Exit Sub
    End If


    Set rg_asset_nick = mySht.Range("rng_AIR_LayerName")
    Set rg_analysis_id = mySht.Range("rng_AIR_LayerGroup")

    ' Loops over all the cells in the range
    count = 0
    For i = 1 To rg_asset_nick.Rows.count
        nickName = rg_asset_nick.Cells(i, 1)
        analysisId = rg_analysis_id.Cells(i, 1)
        ' Skips void cells
        If nickName <> "" Then
            If (modDBInterface.checkStringKeyExists("tblAsset", "strNick", nickName) = False) And (nickName <> "") Then
                ' Checks that the provided Nick is a valid one
                rg_asset_nick.Cells(i, 1).Select
                MsgBox "Invalid asset Nick <" & nickName & ">"
            ElseIf modDBInterface.checkNumKeyExists("tblCondition", "intId", analysisId) = False Then
                ' Checks that the provided AnalysisID is a valid one
                rg_analysis_id.Cells(i, 1).Select
                MsgBox "Invalid analysis ID <" & analysisId & ">"
            Else
                ' In case no error occurs calls the routine to perform the association
                Call modDBInterface.updateNumValueStringKey("tblAsset", "intCondition", analysisId, "strNick", nickName)
                count = count + 1
            End If
        End If
    Next
    
    ' Print a message
    MsgBox "Number of associations computed: " & count

End Sub
'==============================================================================================

