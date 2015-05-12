Attribute VB_Name = "NOAH_sht_modelRMS"
Option Explicit

'##############################################################################################
' ROUTINES, FORMS AND FUNCTIONS RELATIVE TO THE RMS SHEET
'   [1]  Sub_RMS_LoadDBNames        (to be checked)
'   [2]  Sub_RMS_RetrieveData       (to be developed)
'   [3]  Sub_RMS_ImportELT          (to be developed/reviewed)
'   [4]  Sub_RMS_ShowELT            (to be developed)
'   [5]  Sub_RMS_DeleteELT          (to be developed)
'   [6]  Sub_RMS_CalculateRisk      (to be developed)
'   [7]  Sub_RMS_SubmitOEP          (to be developed)
'   [8]  Form_RMS_ImportELT         (to be reviewed)
'   [9]  Form_RMS_ShowELT           (to be reviewed)
'   [10] Form_RMS_DeleteELT         (to be developed)
'##############################################################################################


Public Sub Sub_RMS_LoadDBNames(ByRef myWb As Workbook, _
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
' Descr:   Handles the "Load Companies" button_click event.
'          in the relevant sheet populates the ComboBox containing the names of those
'          Companies which have been created within the local Microsoft SQL Database
'
' Version:
' IL - 17/07/2014
'==============================================================================================
                               
' Set Global Variables
Set wb1 = myWb
Set sh1 = mySht

' Set Local Variables
Dim SQLtype As String
Dim SQLname As String
Dim DBname As String
Dim DBlist As String
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strSQL As String
    
        ' Inizializes values
        SQLtype = "RMS"
        SQLname = mySht.Range("rng_RMS_SQLserver")
        DBname = "master"

        ' Open the connection to the Database
        Set cnn = New ADODB.Connection
        cnn.ConnectionString = strConn(SQLtype, SQLname, DBname)
        cnn.Open

        ' Retrive the available analyses
        strSQL = sqlRMS_AvailProgramme
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cnn, adOpenDynamic, adLockOptimistic
        
        ' Filters only RDM database (i.e. containing ELT)
        'rs.Filter = rs.Fields(0).Name & " LIKE '*RDM*'"
        ' note, some dbs don't have rdm in their name
        
        rs.MoveFirst
        
        ' Clears and Populates the ComboBox
        myCmb.value = ""
        With myCmb
            .Clear
            rs.MoveFirst
            Do While rs.EOF = False
                myCmb.AddItem rs.Fields(0).value
                rs.MoveNext
            Loop
        End With
    
End Sub
'==============================================================================================





Public Sub Sub_RMS_RetrieveData(ByRef myWb As Workbook, _
                                ByRef mySht As Worksheet)
'==============================================================================================
' [2] SUB: TO RETRIEVE STORED ANALYSES
'==============================================================================================
' Input:   #myWb    :  the reference to the workbook which calles this routine
'          #myForm  :  the reference to the form which called this routine
' Output:  -
' Descr:   Handled the "Retrieve Analysis" button_click event.
'          Retrieves all the analyses that have been loaded and associated to the CatBond/CatSwapProgram
'          and lists them into the "Available ELT" table
' Version:
' DB 2015-01-30
'==============================================================================================

Dim numAnalyses As Integer
Dim ownerCode, ownerType As String

Dim rg1, rg2, rg3, rg4 As Range
Dim rs As Recordset

On Error GoTo errFunction

If checkSchedaType(myWb) <> WS_CATSWAP Then
    MsgBox "Error, invalid SCHEDA_TYPE, exit", vbCritical
    End
End If
    
    ' get analysis owner data
    ownerCode = getOwnerCode(myWb)
    ownerType = getOwnerType(myWb)

numAnalyses = modDBInterface_RMSData.getAnalysesCount_byUMR(ownerCode)
'MsgBox "" & numAnalyses & " analyses found for <" & ownerCode & "," & ownerType & " >"


' add rows if needed
On Error GoTo errFunction
Set rg1 = Range("rng_RMS_AnalysesID")
Set rg2 = Range("rng_RMS_AnalysesDescr")
Set rg3 = rg2.Cells(rg2.Rows.count - 1, 1)
Set rg4 = mySht.Cells.Range(rg1.Cells(1, 1), rg3)
rg4.Select

'resize(adds) range and clears it
Call mod_rng_resize.range_insertRows_worksheet(rg4, numAnalyses)
rg1.ClearContents
rg2.ClearContents
rg1.Cells(1, 1).Select

    ' get analyses list
Call select_RMSAnalysis_forProgram(ownerCode, rs)
If Not rs Is Nothing Then
    rg4.CopyFromRecordset rs
End If


On Error GoTo 0
Exit Sub
errFunction:
On Error GoTo 0
MsgBox "[NOAH_sht_modelAIR:Sub_RMS_RetrieveData] Error: " & Err.Description



End Sub
'==============================================================================================



Public Sub Sub_RMS_ImportELT(ByRef myWb As Workbook, _
                             ByRef mySht As Worksheet, _
                             ByRef myForm As Object)
'==============================================================================================
' [3] SUB: TO IMPORT ELT
'==============================================================================================
' Input:   #myWb    :  the reference to the workbook which calles this routine
'          #myForm  :  the reference to the form which called this routine
'==============================================================================================
' IT IMPORT INTO THE MySQL DATABASE THE ELT ASSOCIATED TO THE SELECTED ANALYSIS
'----------------------------------------------------------------------------------------------
' Input:   #myWb    :  the reference to the workbook which calles this routine
'          #myForm  :  the reference to the form which called this routine
' Output:  -
' Descr:    Handles the "SUBMIT" button_click event in the "Import ELT" form.
'           This Sub does the main work, reads data from SQL Server and puts it to MySQL.
'           Generates and imports the ELT into MySQL.
'           The user can COMPOSE multiple analyses from the orginal RMS RDM database,
'           setting "1" in the "Include" column in the UserForm.
'           PERIL and REGION can be overridden using the UserForm.
' Version:
' DB 2015-01-30
'----------------------------------------------------------------------------------------------
Dim cnn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim idELT As Long
Dim perspcode, fileName As String
Dim obj

Dim peril As String, region As String, sql As String
Dim analysisId As Integer, i As Integer
    
    'if analysis description already exists, exit and choose another one
    If checkStringKeyExists("tblRMSList", "strGroupName", myForm.txt_AnalysisDescription) Then
        MsgBox "La Description selezionata, <" & myForm.txt_AnalysisDescription & "> è già presente nel DB, sceglierne un'altra."
        Exit Sub
    End If
    
    ' inserts a new (empty) RMS ELT, and get the ID into idELT
    Call insert_RMSAnalysis(myForm.txt_AnalysisDescription, idELT)
    ' associate the new ELT to the current catswapProgram
    Call assoc_CatSwapProgram_RMSanalysis(myWb.Worksheets("Summary").Range("rng_UMR"), idELT)
    
    ' connects to SQL Server
    Set cnn = getSQLServerConnectionSSPI(myForm.txt_SQLname)
    ' uses the selected db
    cnn.Execute ("USE " & myForm.txt_CompanyName)
    
    ' gets the selected PERSPVALUE
    For Each obj In myForm.Controls
        If TypeName(obj) = "OptionButton" Then
            If obj.value = True Then
                perspcode = obj.Caption
            End If
        End If
    Next

    ' set the temporary output CSV filename
    fileName = myWb.Path & "\" & "ELT_RMS_" & idELT & ".csv"
    
    '#################################################################################
    '## generate the ELT from SQL Server and read it
    '## use HELPER object
    Dim oRMSh As cRMS_helper
    Set oRMSh = New cRMS_helper
    ' init the helper
    Call oRMSh.init
    ' must set these parameters
    oRMSh.intRMS = idELT
    oRMSh.perspcode = perspcode
    
    ' CYCLE for each row selected in the Submit form
    With myForm.tbl_Analyses
        i = 1
        'while table is not empty
        Do While .Columns.Cells(i, 2).value <> ""
            ' if analysis is selected for import
            If .Columns.Cells(i, 1).value <> "" Then
                'gets data from the selected row
                peril = .Columns.Cells(i, 5).value
                region = .Columns.Cells(i, 6).value
                analysisId = .Columns.Cells(i, 2).value
                
                ' puts data to the helper
                Call oRMSh.addAnalysis(analysisId, region, peril)
                
            End If
            i = i + 1
        Loop
    End With
    
    '################################################
    '## generate Temp table in SQL Server
    Call oRMSh.generateCombinedAnalysis(cnn)
    
    ' gets the data from the temp table to the recordset
    Call oRMSh.retrieveCombinedData(cnn, rs)
    
    ' rs could be empty! Print and exit
    If check_Recordset_NotEmpty(rs) = False Then
        Call MsgBox("[NOAH_sht_modelRMS:Sub_RMS_ImportELT] Warning, the ELT from RMS is empty, no points are imported for this ELT <" & myForm.txt_AnalysisDescription & "> with ID <" & idELT & ">", vbExclamation)
        Exit Sub
    End If
           
    ' writes the recordset to file
    Call ExportToCsvELTRMS(rs, fileName)

    ' copy the exported CSV to CSV_repo
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFile(fileName, "K:\Shared Modeling\ANALISI FUNZIONALI\NOAH\Database\CSV_repo\ELT_RMS_" & idELT & ".csv")
    
    ' gets loading command
    sql = getSQLloadCsvELT(fileName)
    ' loads the CSV into the DB
    Call DB_Utilities.execCommandSQL(sql)
    ' delete the temporary csv
    Call RemoveFile(fileName)
    
    MsgBox Prompt:="Work done, the ELT has been imported with idELT = " & idELT

End Sub
'==============================================================================================



Public Sub Sub_RMS_CalculateRisk(ByRef myWb As Workbook, ByRef mySht As Worksheet)
'==============================================================================================
' [6] SUB: TO CALCULATE RISK MEASURES
'==============================================================================================
' Descr:   Handles the "Estimate Risk" button_click event.
'          Demo version
' Version:
' DB 2015-01-30
'=============================================================================================
Set wb1 = myWb
Set sh1 = mySht

Dim rg_ELAgg As Range, rg As Range
Dim valOut, layerName As String
Dim i As Integer

Set rg = sh1.Range("rng_RMS_LayerName")
Set rg_ELAgg = sh1.Range("rng_RMS_RiskMeasure_ELaggregate")

' refresh data in the "Risk Measures" table
Call getAssetList_RMS

If rinterface.ServerIsConnected = False Then
    Call rinterface.StartRServer
End If

For i = 1 To rg.Rows.count
    layerName = rg.Cells(i, 1).value
    If layerName = "" Then GoTo nextIteration
    
    ' puts layer Name
    Call rinterface.PutArrayFromVBA("layerName", layerName)
    ' gets ELT into R
    Call rinterface.RunRFile("K:\Shared Modeling\ANALISI FUNZIONALI\NOAH\Codes\R\RMS_ELT_getData_fromLayerName.r")
    ' init Variables
    Call rinterface.RunRFile("K:\Shared Modeling\ANALISI FUNZIONALI\NOAH\Codes\R\Demo Code\load_elt_init2.r")
    ' computes EL
    Call rinterface.RunRFile("K:\Shared Modeling\ANALISI FUNZIONALI\NOAH\Codes\R\Demo Code\code2.r")
    
    ' retrieves EL
    valOut = rinterface.GetArrayToVBA("el")
    
    ' prints out EL
    rg_ELAgg.Cells(i, 1).value = valOut(0, 0)
    ' formats cell
    rg_ELAgg.Cells(i, 1).NumberFormat = "0.000%"
    
nextIteration:
Next

End Sub
'==============================================================================================



Public Sub Sub_RMS_SubmitOEP(ByRef myWb As Workbook, ByRef mySht As Worksheet)
'==============================================================================================
' Descr:   Handles the "Submit OEP"/"Work with ELTs" button_click Event.
'          Shows the UserForm to associate Assets with ELT/Analyses.
'==============================================================================================
    Set wb1 = myWb
    Set sh1 = mySht
    userForm_RMS_Analysis.Show
End Sub
'==============================================================================================



Public Sub Form_RMS_ImportELT(ByRef myWb As Workbook, ByRef mySht As Worksheet)
'==============================================================================================
' [8] FORM: TO IMPORT ELT
'==============================================================================================
' Input:   #myWb    :  the reference to the workbook which calles this routine
'          #myForm  :  the reference to the form which called this routine
' Output:  -
' Descr:   Handles the "Import ELT" button_click event. Initializes the form to import ELTs from RMS
'          and shows it.
' Version:
' IL - 17/07/2014
'==============================================================================================

    ' Sets Global Variables
    Set wb1 = myWb
    Set sh1 = mySht
    
    ' Defines Local Variable
    Dim myForm As Object
    Dim SQLtype As String
    Dim SQLname As String
    Dim DBname As String
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    
        ' Inizializes values
        Set myForm = userForm_RMS_ImportELT
        SQLtype = "RMS"
        SQLname = mySht.Range("rng_RMS_SQLserver")
        DBname = mySht.Range("rng_RMS_CompanyList")

        ' Open the connection to the Database
        Set cnn = New ADODB.Connection
        cnn.ConnectionString = strConn(SQLtype, SQLname, DBname)
        cnn.Open

        ' Retrive the available analyses
        strSQL = sqlRMS_AvailAnalyses
        Set rs = New ADODB.Recordset
        rs.Open strSQL, cnn, adOpenDynamic, adLockOptimistic

        ' Load the form to import ELT
        Load myForm
        
        ' Initialize values of objects on the form
        myForm.txt_SQLname = SQLname
        myForm.txt_CompanyName = DBname
        myForm.opt_RL.value = True
        With myForm.tbl_Analyses
            i = 1
            .Cells.ClearContents
            Do While rs.EOF = False
                .Cells(i, 2) = rs.Fields("ID").value
                .Cells(i, 3) = rs.Fields("NAME").value
                .Cells(i, 4) = rs.Fields("DESCRIPTION").value
                .Cells(i, 5) = rs.Fields("PERIL").value
                .Cells(i, 6) = rs.Fields("REGION").value
                i = i + 1
                rs.MoveNext
            Loop
            .Cells(1, 1).Select
        End With

        myForm.Show
        
End Sub
'==============================================================================================



Public Sub Form_RMS_ShowELT(ByRef myWb As Workbook, ByRef mySht As Worksheet)
'==============================================================================================
' [9] FORM: TO SELECT THE ELT TO BE IMPORTED
'==============================================================================================
' Input:    #myWb   :  the workbook which is calling the routine
'           #mySht  :  the worksheet which is calling the routine
' Output:   -
' Descr:    Handles the "Show ELT" button_click event. Shows the selected ELT into a new sheet
' Vers:
' DB 2015-01-30
'----------------------------------------------------------------------------------------------


Dim rg, rg2, rg3 As Range
Dim intID As Long
Dim newWs As Worksheet
Dim rs As Recordset

'On Error GoTo errFunction
Set rg = Range("rng_RMS_AnalysesID")

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
Call DB_Utilities.execTableSQL_withRS("select * from tbleltrms where intrms=" & intID & " order by inteventid;", rs)
Call writeRecordsetColumnHeader(newWs.Cells(1, 2), rs)
Call newWs.Cells(2, 2).CopyFromRecordset(rs)

MsgBox "Done"
On Error GoTo 0
Exit Sub
errFunction:
On Error GoTo 0
MsgBox "[NOAH_sht_modelAIR:Form_RMS_ShowELT] Error: " & Err.Description



End Sub
'==============================================================================================



Public Sub Form_RMS_DeleteELT(ByRef myWb As Workbook, _
                              ByRef mySht As Worksheet)
'==============================================================================================
' [10] FORM: TO DELETE STORED ELT
'==============================================================================================
' Input:    #myWb   :  the workbook which is calling the routine
'           #mySht  :  the worksheet which is calling the routine
' Output:   -
' Descr:    Handles the "Delete ELT" button_click event. Deletes the selected ELT
' Vers:
' DB 2015-01-30
'==============================================================================================

Dim rg, rg2, rg3 As Range
Dim intID As Long

'On Error GoTo errFunction
Set rg = Range("rng_RMS_AnalysesID")

Set rg2 = ActiveCell.Cells(1, 1)

Set rg3 = mySht.Cells(rg2.Row, rg.column)
rg3.Select

If Not mod_Checks.checkSingleRangeType(rg3, "INT") Or (rg3.value = "") Then
    MsgBox "Invalid ELT Id found <" & rg3.value & "> , exit."
    End
End If

'asks the user
If (MsgBox("Delete the ELT with ID <" & rg3.value & "> ?", vbYesNo) = vbYes) Then
    intID = rg3.value
    Call modDBInterface.deleteNumKey("tblRMSList", "intId", intID)
    rg3.Cells(1, 1).ClearContents
    MsgBox "Deleted"
End If

On Error GoTo 0
Exit Sub
errFunction:
On Error GoTo 0
MsgBox "[NOAH_sht_modelAIR:Form_RMS_DeleteELT] Error: " & Err.Description

End Sub
'==============================================================================================


' clears the "Risk Measures" table and then fills it
Public Sub getAssetList_RMS()
    Dim s, strSQL, res
    Dim rg As Range, rg1 As Range
    
    s = mod_helper.getOwnerCode(wb1)
    
    'clears output table
    Call clearRiskMeasuresRms
    
    Set rg = sh1.Range("rng_RMS_LayerName")
    Set rg1 = sh1.Range("rng_RMS_LayerStructure")
    
    ' gets asset list
    strSQL = "SELECT strLayerName, intRmsAnalysisToProgram from tblcatswaplayer where strProgramUMR='" & s & "' order by intlayernum;"
    res = mod_btn_GetData.putRsIntoRange_addByRow(strSQL, rg)
    
    If (res = 0) Then
        Call MsgBox("No assets are associated to this Worksheet/CatSwapProgram/CatBond." & vbCrLf & _
            "Swap Analysis/Security Analysis has not been submitted yet?" & vbCrLf & vbCrLf & _
            "Please click <Submit Data> in the <Summary> worksheet first.", vbInformation)
    End If
    
    'gets limit and deductibles
    strSQL = "SELECT dblLimit, dblDeductible from tblcatswaplayer where strProgramUMR='" & s & "' order by intlayernum;"
    res = mod_btn_GetData.putRsIntoRange_addByRow(strSQL, rg1)
End Sub

' clears the "Risk Measures" table
Public Sub clearRiskMeasuresRms()
    
    sh1.Range("rng_RMS_LayerName").ClearContents
    sh1.Range("rng_RMS_LayerGroup").ClearContents
    sh1.Range("rng_RMS_LayerStructure").ClearContents
    sh1.Range("rng_RMS_RiskMeasure").ClearContents
    sh1.Range("rng_RMS_RiskAmount").ClearContents
    
    ' update buttons first!
    sh1.Range("rng_RMS_RiskMeasure_ELaggregate").ClearContents

End Sub


