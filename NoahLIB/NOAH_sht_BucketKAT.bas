Attribute VB_Name = "NOAH_sht_BucketKAT"
Option Explicit

'##############################################################################################
' ROUTINES, FORMS AND FUNCTIONS RELATIVE TO THE BUCKET SHEET
'   [1]  Sub_BucketKAT_SubmitData          (to be developed)
'   [2]  Sub_BucketKAT_RetrieveData        (to be developed)
'   [3]  Sub_BucketKAT_DeleteData          (to be developed)
'##############################################################################################


Public Sub Sub_BucketKAT_SubmitData(ByRef myWb As Workbook, _
                                    ByRef mySht As Worksheet)
'==============================================================================================
' [1] SUB: TO SUBMIT KATARSIS BUCKET INTO MySQL DB
'==============================================================================================

 

Dim ownerCode, assetNick, ownerType, assetCode As String

Dim rgBuckets, rgNames, rgValues As Range
Dim rs As Recordset
Dim i, j, n, numCols As Long
Dim contribValue As Double

On Error GoTo errFunction


If checkSchedaType(myWb) = WS_ERROR Then
    MsgBox "Error, unable to find SCHEDA_TYPE, exit", vbCritical
    End
End If
    
    ' get analysis owner data
    ownerCode = getOwnerCode(myWb)
    ownerType = getOwnerType(myWb)

Set rgBuckets = Range("rng_Buckets")
Set rgNames = rgBuckets.Columns(1)

If (ownerType = "CB") Then

    Set rgValues = rgBuckets.Columns(2)
    rgValues.Select
    If MsgBox("Set these buckets for the Cat Bond <" & ownerCode & "> ?", vbYesNo) = vbYes Then
    
        ' delete buckets for this asset
        Call modDBInterface.deleteStringKey("tblAssetBucket", "strAssetCode", ownerCode)
        
        n = rgNames.Rows.count
        For i = 1 To n
            contribValue = rgValues.Cells(i, 1).value
            If (contribValue > 0) Then
                Call insert_Asset_BucketKatarsis(ownerCode, rgNames.Cells(i, 1), contribValue)
            End If
        Next
        MsgBox "Update finished"
    Else
        MsgBox "Buckets update canceled, exit.", vbInformation
        End
    End If
    
End If


If (ownerType = "RE") Then
    numCols = rgBuckets.Columns.count
    
    For j = 2 To numCols
        rgBuckets.Cells(-3, j).Select
        assetNick = rgBuckets.Cells(-3, j).value
        
        
        If assetNick = "" Or (modDBInterface.checkStringKeyExists("tblasset", "strNick", assetNick) = False) Then
            MsgBox "Invalid asset Nick <" & assetNick & "> , exit."
            End
        End If
        
        
        Set rgValues = rgBuckets.Columns(j)
        rgValues.Select
        If MsgBox("Set these buckets for the asset <" & assetNick & "> ?", vbYesNo) = vbYes Then
            
            assetCode = get_AssetCode_byNick(assetNick)
            ' delete buckets for this asset
            Call modDBInterface.deleteStringKey("tblAssetBucket", "strAssetCode", assetCode)
            
            n = rgNames.Rows.count
            For i = 1 To n
                contribValue = rgValues.Cells(i, 1).value
                If (contribValue > 0) Then
                    Call insert_Asset_BucketKatarsis(assetCode, rgNames.Cells(i, 1), contribValue)
'                        Call insert_Asset_numBucketKatarsis(assetCode, i, contribValue)
                End If
            Next
        End If
        
    Next
    
    MsgBox "Update finished"
    
End If


On Error GoTo 0
Exit Sub
errFunction:
On Error GoTo 0
MsgBox "[NOAH_sht_modelAIR:Sub_BucketKAT_SubmitData] Error: " & Err.Description

End Sub
'==============================================================================================



Public Sub Sub_BucketKAT_RetrieveData(ByRef myWb As Workbook, _
                                      ByRef mySht As Worksheet)
'==============================================================================================
' [2] SUB: TO RETRIEVE KATARSIS BUCKET FROM MySQL DB
'==============================================================================================

    MsgBox "Routine 'Sub_BucketKAT_RetrieveData' is Under Construction "

End Sub
'==============================================================================================



Public Sub Sub_BucketKAT_DeleteData(ByRef myWb As Workbook, _
                                    ByRef mySht As Worksheet)
'==============================================================================================
' [3] SUB: TO DELETE KATARSIS BUCKET FROM THE MySQL DB
'==============================================================================================

    MsgBox "Routine 'Sub_BucketKAT_DeleteData' is Under Construction"

End Sub
'==============================================================================================
