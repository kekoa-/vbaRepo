Attribute VB_Name = "NOAH_sht_BucketHR"
'##############################################################################################
' ROUTINES, FORMS AND FUNCTIONS RELATIVE TO THE HR SHEET
'   [1]  Form_BucketHR_SelectBucket             (to be reviewed)
'   [2]  Sub_SetBucketHR                        (to be reviewed)
'   [3]  Sub_BucketHR_SubmitData                (to be developed)
'   [4]  Sub_BucketHR_RetrieveData              (to be developed)
'   [5]  Sub_BucketHR_DeleteData                (to be developed)
'##############################################################################################



Public Sub Form_BucketHR_SelectBucket(ByRef myWb As Workbook, _
                                      ByRef mySht As Worksheet)
'==============================================================================================
' [1] FORM LOAD: TO SELECT THE HR BUCKETS
'==============================================================================================
' Input:    #myWb   :  the workbook which is calling the routine
'           #mySht  :  the worksheet which is calling the routine
' Output:   -
' Descr:    Shows the form to select the HR Buckets to be assigned to the layers of the
'           programme under analysis
' Vers:
' 30.07.2014 - IL
'----------------------------------------------------------------------------------------------
Dim myForm As Object
Dim rngLayer As Range, rngBucket As Range
Dim list() As String

    ' Set Global Variables
    Set wb1 = myWb
    Set sh1 = mySht
    
    ' Set Local Variables
    Set myForm = userForm_SelectBucket
    Set rngLayer = mySht.Range("rng_HR_Layer")
    
    ' Load the relevant form
    Load myForm
    
    ' Populates the ListBox containing layers' names
    For i = 1 To rngLayer.Rows.count
        With myForm.list_SelectLayer
            If rngLayer.Cells(i, 1) <> "" Then
                .AddItem
                .list(i - 1) = rngLayer.Cells(i, 1).value
            End If
        End With
    Next
    
    ' Populates the ListBox containing buckets' names
    list = Param_ListBucketHR
    For i = 1 To UBound(list)
        With myForm.list_SelectBucket
            .AddItem
            .list(i - 1) = list(i)
        End With
    Next
    
    ' Shows the relevant form
    myForm.Show
    

End Sub
'==============================================================================================



Public Sub Sub_SetBucketHR(ByRef myForm As UserForm)
'==============================================================================================
' [2] SUB CALL: TO SET HR BUCKETS
'==============================================================================================
' Input:    #myForm   :  the the userForm which is calling the routine
' Output:   -
' Descr:    Given the selected layers and buckets within the userForm, it properly
'           populates relevant ranges in the HR sheet
' Vers:
' 30.07.2014 - IL
'==============================================================================================

' Define Local Variables
Dim mySht As Worksheet
Dim i As Long, k As Long, z As Long
    
    ' Identifies the destination sheet
    Set mySht = sh1
    
    ' Loops over the selected layers
    
    With myForm
        ' Checks if the user has selected at least one layer and one bucket
        If .list_SelectLayer.ListCount = 0 Or .list_SelectBucket.ListCount = 0 Then
            MsgBox Prompt:="Attention! No layer or bucket have been selected", _
                   Title:="NOAH - Error"
            Exit Sub
        End If
        
        ' Loops over all the HR Buckets
        z = 0
        For k = 0 To .list_SelectBucket.ListCount - 1
            If .list_SelectBucket.Selected(k) = True Then
                z = z + 1
                ' Associates the selected HR bucket to the relevant Layer
                For i = 0 To .list_SelectLayer.ListCount - 1
                    If .list_SelectLayer.Selected(i) = True Then
                        mySht.Range("rng_HR_Bucket").Cells(i + 1, z) = .list_SelectBucket.list(k)
                    End If
                Next
            End If
        Next
    End With
    
End Sub
'==============================================================================================



Public Sub Sub_BucketHR_SubmitData(ByRef myWb As Workbook, _
                                   ByRef mySht As Worksheet)
'==============================================================================================
' [3] SUB CALL: TO SUBMIT DATA INTO THE MySQL DATABASE
'==============================================================================================

Dim assetNick, bucketName As String
    
Dim rgBuckets, rgNames, rgConflict As Range


Dim i, j, n, numRows, numCols, conflictCode As Long

On Error GoTo errFunction

Set rgBuckets = Range("rng_HR_Bucket")
Set rgNames = Range("rng_HR_Layer")
Set rgConflict = Range("rng_HR_Rating")


numRows = rgNames.Rows.count

For i = 1 To numRows
    assetNick = rgNames.Cells(i, 1).value
    
    ' if empty then skip
    If assetNick = "" Then GoTo nextIteration
    
    ' if invalid then warning and skip
    If (modDBInterface.checkStringKeyExists("tblasset", "strNick", assetNick) = False) Then
        MsgBox "Invalid asset Nick <" & assetNick & "> , skip."
        GoTo nextIteration
    End If
    

    rgBuckets.Rows(i).Select
    If MsgBox("Set these Hannover buckets for the CatSwap layer <" & assetNick & "> ?", vbYesNo) = vbYes Then
        
        ' delete existing buckets for this catwap layer
        Call modDBInterface.deleteStringKey("tblCSLayerHRBucket", "strLayerName", assetNick)
        
        ' insert buckets
        n = rgBuckets.Columns.count
        For j = 1 To n
            bucketName = rgBuckets.Cells(i, j).value
            If (bucketName <> "") Then
                Call insert_CatSwapLayer_HRBucket(assetNick, bucketName)
            End If
        Next
        
        ' updates conflict
        If (rgConflict.Cells(i, 1).value = 0) Or (rgConflict.Cells(i, 1).value = 1) Or (rgConflict.Cells(i, 1).value = 2) Then
            conflictCode = rgConflict.Cells(i, 1).value
            Call update_CatSwapLayer_conflit(assetNick, conflictCode)
        End If
        
    End If
    
    


        
        
nextIteration:
Next

On Error GoTo 0
Exit Sub
errFunction:
On Error GoTo 0
MsgBox "[NOAH_sht_BucketHR:Sub_BucketHR_SubmitData] Error: " & Err.Description

End Sub
'==============================================================================================



Public Sub Sub_BucketHR_RetrieveData(ByRef myWb As Workbook, _
                                     ByRef mySht As Worksheet)
'==============================================================================================
' [4] SUB CALL: TO RETRIEVE DATA FROM THE MySQL DATABASE
'==============================================================================================

    MsgBox "Routine 'Sub_BucketHR_RetrieveData' is Under Construction"

End Sub
'==============================================================================================



Public Sub Sub_BucketHR_DeleteData(ByRef myWb As Workbook, _
                                   ByRef mySht As Worksheet)
'==============================================================================================
' [5] SUB CALL: TO DELETE DATA FROM THE MySQL DATABASE
'==============================================================================================

    
'    Call modDBInterface_Buckets.delete_CatSwapLayer_HRBucket(layerName)
    
    MsgBox "Routine 'Sub_BucketHR_DeleteData' is Under Construction"

End Sub
'==============================================================================================






Public Function Param_ListBucketHR() As String()
'==============================================================================================
' PARAMETER: LIST OF HR BUCKETS
'----------------------------------------------------------------------------------------------
' Input:    -
' Output:   an array of string
' Descr:    return an array of strings containing the HR buckets' names
' Vers:
' 30/07/2014: IL
'----------------------------------------------------------------------------------------------

Dim list(1 To 51) As String

    list(1) = "Africa"
    list(2) = "Asia China / Hkg Nat Cat"
    list(3) = "Asia India"
    list(4) = "Asia JP EQ 1st Event"
    list(5) = "Asia JP Wind"
    list(6) = "Asia Korea"
    list(7) = "Asia Middle East"
    list(8) = "Asia Other Nat Cat"
    list(9) = "Asia Taiwan Nat Cat"
    list(10) = "AUS EQ"
    list(11) = "NZ EQ"
    list(12) = "AUS/NZ Wind"
    list(13) = "Aviation excl. War & Terror"
    list(14) = "Aviation War & Terror"
    list(15) = "Crop US"
    list(16) = "Energy non-elemental"
    list(17) = "EU Eastern Europe & Austria  EQ"
    list(18) = "EU Other EQ"
    list(19) = "EU UK Flood"
    list(20) = "EU Wind - Northern Path 1st event (UK, F, Benelux, DE, Scandinavia, PL)"
    list(21) = "EU Wind - Northern Path 2nd event (UK, F, Benelux, DE, Scandinavia, PL)"
    list(22) = "EU Wind - Southern Path 1st Event (UK, F, Benelux, DE, CZ, AU)"
    list(23) = "EU Wind - Southern Path 2nd event (UK, F, Benelux, DE, CZ, AU)"
    list(24) = "Latin America Nat Cat Carribbean / South America"
    list(25) = "Latin America Nat Cat Carribbean / Florida"
    list(26) = "Latin America Nat Cat Carribbean / Texas"
    list(27) = "Latin America Nat Cat Mexico"
    list(28) = "Latin America Nat Cat Chile"
    list(29) = "Latin America Brazil"
    list(30) = "Latin America Nat Cat other South America"
    list(31) = "North America EQ California "
    list(32) = "North America EQ Pacific NW"
    list(33) = "North America EQ Canada"
    list(34) = "North America EQ New Madrid"
    list(35) = "North America Hurricane US - FL to Texas +GOM + Mexico Wind 1st Event"
    list(36) = "North America Hurricane US - FL to Texas +GOM + Mexico Wind 2nd Event"
    list(37) = "North America Hurricane US - FL-NC 1st Event"
    list(38) = "North America Hurricane US - FL-NC 2nd Event"
    list(39) = "North America Hurricane US - VA-ME "
    list(40) = "North America Hurricane US - VA-ME 2nd event"
    list(41) = "North America Tornado"
    list(42) = "Terror - Denmark"
    list(43) = "Terror - Latin America ex Mexico"
    list(44) = "Terror - Latin America Mexico"
    list(45) = "Terror - UK"
    list(46) = "Terror - US"
    list(47) = "Terror - India"
    list(48) = "Terror - WW"
    list(49) = "Turkey Catastrophe 1st Event"
    list(50) = "Turkey Catastrophe 2ndEvent"
    list(51) = "Rest - no separate bucket"

    Param_ListBucketHR = list

End Function
'==============================================================================================

