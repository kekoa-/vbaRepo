Attribute VB_Name = "modDBInterface_CatBond"
Option Explicit


Sub SubmitDataIntoDatabase_CatBond(ByRef wb As Workbook)
'==============================================================================================
'IT SUBMITS INPUT DATA INTO THE DATABASE
'----------------------------------------------------------------------------------------------
' Input:   #wb  :  the workbook which the the data have to be exported from
' Output:  -
' Descr:
'----------------------------------------------------------------------------------------------
    
    
    Dim sh As Worksheet
    Dim assetCode, nickName As String
    Dim rgcode As Range
    
    
    If MsgBox("Update di tutte le informazioni nel DB per questo Cat Bond, procedere?", vbOKCancel) <> vbOK Then
        MsgBox ("Update Canceled")
        Exit Sub
    End If
    
    
    Set sh = wb.Worksheets("Summary")
    assetCode = sh.Range("rng_strasset_code").value
    nickName = sh.Range("rng_strasset_nick").value
    
    Set rgcode = sh.Range("rng_strasset_code")
    rgcode.Interior.PatternColorIndex = xlAutomatic
    rgcode.Interior.Color = 65535
    rgcode.Select
    
    ' checks for invalid characters
    Call check_ASCIIcodeForKey_intoRange_orExit(sh.Range("rng_strasset_code"))
    Call check_ASCIIcodeForKey_intoRange_orExit(sh.Range("rng_strasset_nick"))
    
    
    
    ' check if the NICK for this CatBond is already in the DB.
    ' if it is, then we need to update the AssetCode
    If checkStringKeyExists("tblAsset", "strNick", nickName) Then
        If MsgBox("Il CatBond con Nick <" & nickName & "> è già presente nel DB, modificare l'ID per questo asset a: <" & assetCode & "> ?", vbYesNo) = vbYes Then
            ' update the key
            Call execCommandSQL(" select intid into @assetId from tblasset where strNick='" & nickName & "' ;")
            
            If checkStringKeyExists("tblAsset", "strCode", assetCode) Then
                MsgBox "Errore, l'ID <" & assetCode & "> è già assegnato ad un altro asset, update annullato. Exit"
                Exit Sub
            End If
            
            Call execCommandSQL(" update tblasset set strCode ='" & assetCode & "' where intid=@assetId ; ")
        Else
        ' if no, cancel
            MsgBox "Update annullato. Exit"
            rgcode.Interior.Pattern = xlNone
            Exit Sub
        End If
    End If
    
    
    
    ' check if key exists in table

    If checkStringKeyExists("tblAsset", "strCode", assetCode) Then
        If MsgBox("Il CatBond è già presente nel DB, fare update delle informazioni per il CatBond <" & assetCode & "> ?", vbYesNo) = vbYes Then
            ' continue outside of this block..
        Else
            ' if no, cancel
            MsgBox "Update canceled. Exit"
            rgcode.Interior.Pattern = xlNone
            Exit Sub
            
        End If
    
    Else
    ' if doesn not exist yet,
    
        ' ask to create new record
        If MsgBox("Il CatBond non è presente nel DB, inserire CatBond con ID <" & assetCode & "> ?", vbYesNo) = vbYes Then
            ' if yes, create new record
            
            ' insert KEY into tblAsset first
            Call insertStringKey("tblAsset", "strCode", assetCode)
            ' insert NICK into tblAsset
            Call updateStringValueStringKey("tblAsset", "strNick", nickName, "strCode", assetCode)
            'writes "CB" as assetType in tblAsset
            Call modDBInterface.updateStringValueStringKey("tblAsset", "strAssetType", "CB", "strCode", assetCode)
            ' INSERT RECORD into tblCatBondInfo
            Call insertStringKey("tblCatBondInfo", "strAssetCode", assetCode)
        Else
            ' if no, cancel
            MsgBox "Update canceled"
            rgcode.Interior.Pattern = xlNone
            Exit Sub
        End If
    End If
    
    ' non dovrebbe mai entrare qui dentro, cmq --
    If Not checkStringKeyExists("tblcatbondinfo", "strAssetCode", assetCode) Then
        MsgBox "Error, check the log file", vbExclamation
        logToFile ("Err: Creating CatBond info for <" & assetCode & "> , asset exist, catbondinfo does not exist.")
        Call insertStringKey("tblCatBondInfo", "strAssetCode", assetCode)
    End If
    
    ' BEGIN
    ' UPDATE VALUES
    
    ' retrieve field list
    
    'helper objects
    Dim oLoader As cLinkLoader
    Set oLoader = New cLinkLoader
    
    Dim oList As cLinkList
    Set oList = New cLinkList

    ' LOAD field list
    Call oLoader.loadCBFieldsList(oList)

    ' ######### updater
    Dim oUpdater As cUpdater
    Set oUpdater = New cUpdater
    Set oUpdater.thisWb = ActiveWorkbook
    
    
    '######### CHECK ranges
    
    If oUpdater.checkRanges(oList, True) = False Then
        'if check failed
        If MsgBox("Some Ranges are missing in the Excel worksheet, update anyway?", vbYesNo) <> vbYes Then
            MsgBox "Update canceled"
            rgcode.Interior.Pattern = xlNone
            End
        End If
    End If
        
   
    ' here, proceed with update
    Call oUpdater.updateXlsToDB(oList)
    
    ' additional update-- for paymentDates
    Call updateCatbond_additional(wb)
        
    ' updates file path
    Call update_CatBond_XlsPath(assetCode, Replace(wb.FullName, "\", "/"))

    'update field list
    
    MsgBox "Update Finished"
    rgcode.Interior.Pattern = xlNone
    
End Sub
    
    
    
Sub DeleteFromDatabase_CatBond(ByRef wb As Workbook)

    Dim sh As Worksheet
    Dim assetCode As String
    
    
    If MsgBox("This will DELETE this asset from the DB, and all its associated information. Proceed?", vbOKCancel) <> vbOK Then
        MsgBox ("Deletion Canceled")
        Exit Sub
    End If
    
    Set sh = wb.Worksheets("Summary")
    assetCode = sh.Range("rng_strasset_code").value
        
    ' executes delete
    Call deleteStringKey("tblAsset", "strCode", assetCode)
    
    MsgBox "Deleted"
    
End Sub

    
    
Sub updateCatbond_additional(ByRef wb As Workbook)
    Dim rg, rg2 As Range
    Dim dateIn As Date
    Dim strKey As String
    
    Set rg = wb.Worksheets("Additional_Info").Range("rng_datPaymentDates")
    strKey = wb.Worksheets("Summary").Range("rng_strAsset_Code")
    

    Set rg2 = rg.Cells(1, 1)
    If mod_Checks.checkSingleRangeType(rg2, "STR") Then Call modDBInterface.updateStringValueStringKey("tblCatBondInfo", "strPaymentDates1", rg2.value, "strAssetCode", strKey)
    Set rg2 = rg.Cells(2, 1)
    If mod_Checks.checkSingleRangeType(rg2, "STR") Then Call modDBInterface.updateStringValueStringKey("tblCatBondInfo", "strPaymentDates2", rg2.value, "strAssetCode", strKey)
    Set rg2 = rg.Cells(3, 1)
    If mod_Checks.checkSingleRangeType(rg2, "STR") Then Call modDBInterface.updateStringValueStringKey("tblCatBondInfo", "strPaymentDates3", rg2.value, "strAssetCode", strKey)
    Set rg2 = rg.Cells(4, 1)
    If mod_Checks.checkSingleRangeType(rg2, "STR") Then Call modDBInterface.updateStringValueStringKey("tblCatBondInfo", "strPaymentDates4", rg2.value, "strAssetCode", strKey)
        


End Sub





'fa l'update del XLS FilePath
Public Sub update_CatBond_XlsPath(ByVal key As String, ByVal inputVal As String)
    Call updateStringValueStringKey("tblCatBondInfo", "strXlsPath", inputVal, "strAssetCode", key)
End Sub


