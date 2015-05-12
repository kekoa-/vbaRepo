Attribute VB_Name = "modDBInterface_CatSwapProgram"
Option Explicit
' le funzioni seguenti si occupano di fare l'update per i catSwapPrograms
' (tabella tblCatSwapProgram)


' crea un nuovo record nella tabella dei CatSwapPrograms
'Public Sub insert_CatSwapProgram(ByVal strUMR As String)
'    Call insertStringKey("tblCatSwapProgram", "strUMR", strUMR)
'End Sub

'fa DELETE del program/record, in base alla key (strUMR) specificata
'Public Sub delete_CatSwapProgram(strUMR As String)
'    Call deleteStringKey("tblCatSwapProgram", "strUMR", strUMR)
'End Sub




'fa l'update del XLS FilePath
Public Sub update_CatSwapProgram_XlsPath(ByVal strUMR As String, ByVal inputVal As String)
    Call updateStringValueStringKey("tblCatSwapProgram", "strXlsPath", inputVal, "strUMR", strUMR)
End Sub



Sub SubmitDataIntoDatabase_CatSwapProgram(ByRef wb As Workbook)
'==============================================================================================
'IT SUMBIT INPUT DATA INTO THE DATABSE
'----------------------------------------------------------------------------------------------
' Input:   #wb  :  the workbook which the the data have to be exported from
' Output:  -
' Descr: Handles the "Submit Data" button_click event for CatSwaps.
'        Reads data from the Worksheet, performs checks, and writes to MySQL
' Vers : DB 30/01/2015
'----------------------------------------------------------------------------------------------
    
    Dim UMR As String, layerName As String, assetCode As String, newAssetCode As String, Ccy As String
    Dim r As Range
    Dim rgLayers As Range
    Dim i As Long
    
    If MsgBox("This will update all the information in this workbook and replace the info on the DB, proceed?", _
    vbOKCancel, ".:: Cat Swap Update ::.") <> vbOK Then
        MsgBox ("Update Canceled")
        Exit Sub
    End If
    
    
    Dim sh As Worksheet
    Set sh = wb.Worksheets("Summary")
    
    ' checks for invalid characters
    Call check_ASCIIcodeForKey_intoRange_orExit(sh.Range("rng_UMR"))
    Call check_ASCIIcodeForKey_intoRange_orExit(sh.Range("rng_Nick"))
    
    
    UMR = sh.Range("rng_UMR").value
    Ccy = sh.Range("rng_Currency").value
    
    
    '##############################################################
    ' check if Program Exists
    If Not checkStringKeyExists("tblCatSwapProgram", "strUMR", UMR) Then

        ' ask to create new record
        If MsgBox("Il CatSwap Program non è presente nel DB, inserire Program con UMR <" & UMR & "> ?", vbYesNo) = vbYes Then
            ' if yes, create new record
            
            ' INSERT into tblCatSwapProgram
            'Call insert_CatSwapProgram(UMR)
            Call insertStringKey("tblCatSwapProgram", "strUMR", UMR)
        Else
            ' if no, cancel
            MsgBox "Update canceled, exit"
            End
        End If
    End If
    
       
    '##############################################################
    ' for each layer
    i = 1
    
    Set rgLayers = sh.Range("rng_Layer_Name")
    rgLayers.Interior.PatternColorIndex = xlAutomatic
    rgLayers.Interior.Color = 65535
    rgLayers.Select
    
    ' CYCLE over layers
    For Each r In rgLayers
        'check if Layers Exist
        layerName = r.value
        ' this is the asset code
        assetCode = UMR & "_L" & i
        
        If layerName = "" Then
            GoTo nextIteration
        End If
        
        ' if finds invalid character codes, exit
        Call check_ASCIIcodeForKey_intoRange_orExit(r)
        
        
        
        '########################################################################
        '########################################################################
        ' check ASSET first
        
        ' if both already exist, then we are fine, do just one check..
        If (checkStringKeyExists("tblAsset", "strNick", layerName) = True) And _
           (checkStringKeyExists("tblAsset", "strCode", assetCode) = True) Then
           
            If getScalarStringKey("tblAsset", "strNick", "strCode", assetCode) <> layerName Then
                ' this is VERY screwed up!
                Call logToFile("Check this! [tblAsset:strNick=<" & layerName & ">]" & _
                               "[tblAsset: strCode=<" & assetCode & ">] . ")
                'Call MsgBox("Error, there exist TWO asset, one asset has strNick=<" & layerName & "> and the other asset has strCode=<" & assetCode & "> ." & vbCrLf & _
                '            "To which asset should we associate the CatSwapLayer? Please check, exit.", vbCritical)
                'Exit Sub
                'umm.. associate by nickname
                newAssetCode = getScalarStringKey("tblAsset", "strCode", "strNick", layerName)
                ' use the newAssetCode
                assetCode = newAssetCode
            End If
'            Call updateStringValueStringKey("tblAsset", "strNick", layerName, "strCode", assetCode)
        End If
        
        ' if both do not exist, then we are fine, just insert into tblAsset first
        If (checkStringKeyExists("tblAsset", "strNick", layerName) = False) And _
           (checkStringKeyExists("tblAsset", "strCode", assetCode) = False) Then
            ' insert assetCode
            Call insertStringKey("tblAsset", "strCode", assetCode)
            ' insert the nick=layerName
            Call updateStringValueStringKey("tblAsset", "strNick", layerName, "strCode", assetCode)
        End If
        
        ' if the asset with code <assetCode> does not exist but an asset with NICK <layerName> already exists then
        ' >> just take the assetCode that already is in the table
        If (checkStringKeyExists("tblAsset", "strNick", layerName) = True) And _
           (checkStringKeyExists("tblAsset", "strCode", assetCode) = False) Then
            ' get existing asset Code
            newAssetCode = getScalarStringKey("tblAsset", "strCode", "strNick", layerName)
            ' notify the user
            'Call MsgBox("Note: assetCode <" & assetCode & "> does not exist in tblAsset, " & _
            '            " but the Nick Name <" & layerName & "> already exists in tblAsset, with asset Code <" & _
            '            newAssetCode & ">. " & " Using the esisting assetCode <" & newAssetCode & "> for layer <" & _
            '            layerName & "> .", vbExclamation)
            'log
            Call logToFile("Note: assetCode <" & assetCode & "> does not exist in tblAsset, " & _
                            " but the Nick Name <" & layerName & "> already exists in tblAsset, with asset Code <" & _
                            newAssetCode & ">. " & " Use the esisting assetCode <" & newAssetCode & "> for layer <" & _
                            layerName & "> .")
            ' use the newAssetCode
            assetCode = newAssetCode
        End If
        
        
        ' if the asset with code <assetCode> does exist but an asset with NICK <layerName> DOES NOT exists then..
        ' >> use a MODIFIED assetCode
        ' >> INSERT new asset with this code
        If (checkStringKeyExists("tblAsset", "strNick", layerName) = False) And _
           (checkStringKeyExists("tblAsset", "strCode", assetCode) = True) Then
            'overwrite layerName
            newAssetCode = assetCode & "_tempcode_" & CInt(Rnd() * 10000)
            assetCode = newAssetCode
            ' insert assetCode
            Call insertStringKey("tblAsset", "strCode", assetCode)
            ' insert the nick=layerName
            Call updateStringValueStringKey("tblAsset", "strNick", layerName, "strCode", assetCode)
        End If

        
        ' END check ASSET
        '########################################################################
        '########################################################################
        

        Call updateStringValueStringKey("tblAsset", "strAssetType", "RE", "strCode", assetCode)

        
        
        ' check if catswap layer exists already
        If Not checkStringKeyExists("tblCatSwapLayer", "strLayerName", layerName) Then
        ' if does not exist, create it..
        
            ' notifies the creation of the new record
            Call MsgBox("Il CatSwap Layer non è presente nel DB, inserisco Layer con " & vbCrLf & _
                        "Layer Name: <" & layerName & "> " & vbCrLf & _
                        "Asset Code: <" & assetCode & "> .", vbInformation)
            
            ' create new record for the CatSwap Layer
            Call insertStringKey("tblCatSwapLayer", "strLayerName", layerName)
            
            'updates CODE for the catSwapLayer
            Call updateStringValueStringKey("tblCatSwapLayer", "strCode", assetCode, "strLayerName", layerName)
            
        End If
                
        'updates umr for the catSwapLayer
        Call updateStringValueStringKey("tblCatSwapLayer", "strProgramUMR", UMR, "strLayerName", layerName)
        
        'updates Asset Name
        Call updateStringValueStringKey("tblAsset", "strName", layerName, "strNick", layerName)
        
        'updates CCY in tblASSET
        Call updateStringValueStringKey("tblAsset", "strCcy", Ccy, "strNick", layerName)
        
        'updates layerNUM for the catSwapLayer
        Call updateNumValueStringKey("tblCatSwapLayer", "intLayerNum", i, "strLayerName", layerName)
        'updates assetNUM for the asset
        Call updateNumValueStringKey("tblAsset", "intAssetNum", i, "strNick", layerName)

        'writes "RE" as assetType in tblAsset
        Call modDBInterface.updateStringValueStringKey("tblAsset", "strAssetType", "RE", "strNick", layerName)
        
        
        i = i + 1

nextIteration:
    Next
        
    
    '##############################################################
    ' updates PROGRAM info and LAYERS info,
    
    ' BEGIN
    ' UPDATE VALUES
    
    ' retrieve field list
    
    'helper objects
    Dim oLoader As cLinkLoader
    Set oLoader = New cLinkLoader
    
    Dim oList As cLinkList
    Set oList = New cLinkList

    ' LOAD field list
    Call oLoader.load_CatSwap_FieldsList(oList)

    ' ######### updater
    Dim oUpdater As cUpdater
    Set oUpdater = New cUpdater
    Set oUpdater.thisWb = ActiveWorkbook
    
    
    '######### CHECK ranges
    If oUpdater.checkRanges(oList, False) = False Then
        'if check failed
        If MsgBox("Some Ranges are missing, update anyway?", vbYesNo) <> vbYes Then
            MsgBox "Update canceled"
            End
        End If
    End If
        
    ' here, proceed with update
    Call oUpdater.updateXlsToDB(oList)
    
    ' updates file path
    Call update_CatSwapProgram_XlsPath(UMR, Replace(wb.FullName, "\", "/"))
    
    MsgBox ("Update complete")
    rgLayers.Interior.Pattern = xlNone

End Sub
'==============================================================================================




Sub DeleteFromDatabase_CatSwap_Program(ByRef wb As Workbook)

    Dim sh As Worksheet
    Dim UMR As String
    
    
    If MsgBox("This will DELETE this CatSwap Program from the DB, all the related CatSwap Layers, and all their associated information. Proceed?", vbOKCancel) <> vbOK Then
        MsgBox ("Deletion Canceled")
        Exit Sub
    End If
    
    Set sh = wb.Worksheets("Summary")
    UMR = sh.Range("rng_UMR").value
        
    ' executes delete
    Call deleteStringKey("tblCatSwapProgram", "strUMR", UMR)
    
    MsgBox "Deleted"
    
End Sub


















