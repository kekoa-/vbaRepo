Attribute VB_Name = "modDBInterface_Buckets"
Option Explicit
' le funzioni seguenti si occupano di fare l'update dei Risk buckets


' associa un catSwapLayer con un bucket Hannover:
' inserisce un record nella tabella (tblCSLayerHRBucket)

Public Sub insert_CatSwapLayer_HRBucket(ByVal layerName As String, ByVal bucketName As String)
    Dim strSQL As String
    
    If checkStringKeyExists("tblCatSwapLayer", "strLayerName", layerName) = False Then
        MsgBox "Error: key <" & layerName & "> not found in table tblCatSwapLayer (strLayerName)" & _
                vbCrLf & vbCrLf & "Il Layer <" & layerName & "> non è presente nella tabella layers"
                
        Exit Sub
    End If
    If checkStringKeyExists("tblHRBucket", "strName", bucketName) = False Then
        MsgBox "Error: key <" & bucketName & "> not found in table tblHRBucket (strName)" & _
                vbCrLf & vbCrLf & "Il Bucket <" & bucketName & "> non è presente nella tabella HR Buckets"
        Exit Sub
    End If
    
    strSQL = "INSERT INTO tblCSLayerHRBucket(strLayerName, strHRBucket) VALUES " & _
             "('" & layerName & "','" & bucketName & "')"
    
    Call execCommandSQL(strSQL)
    
End Sub


' rimuove tutti i bucket Hannover associati ad un layer
'Public Sub delete_CatSwapLayer_HRBucket(strLayerName As String, Optional doErrorCheck As Boolean = False)
'    Call deleteStringKey("tblCSLayerHRBucket", "strLayerName", strLayerName, doErrorCheck)
'End Sub

Public Sub update_CatSwapLayer_conflit(ByVal layerName As String, ByVal conflict As Long)
    Call modDBInterface.updateNumValueStringKey("tblcatswaplayer", "intconflict", conflict, "strlayername", layerName)
End Sub










' associa un catSwapLayer con un bucket Katarsis:
' inserisce un record nella tabella (tblAssetBucket)
'Public Sub insert_Asset_numBucketKatarsis(ByVal assetname As String, ByVal bucketId As Long, _
'                    ByVal dblContribution As Double)
'    Dim strSQL As String
'    If checkStringKeyExists("tblAsset", "strCode", assetName) = False Then
'        MsgBox "Error: key <" & assetName & "> not found in table tblAsset (strCode)" & _
'                vbCrLf & vbCrLf & "L'asset <" & assetName & "> non è presente nella tabella asset?"
'
'        Exit Sub
'    End If
'    If checkNumKeyExists("tblBucket", "intId", bucketId) = False Then
'        MsgBox "Error: key <" & bucketId & "> not found in table tblBucket (intID)" & _
'                vbCrLf & vbCrLf & "Il Bucket <" & bucketId & "> non è presente nella tabella Buckets"
'        Exit Sub
'    End If
'    strSQL = "INSERT INTO tblAssetBucket(strAssetCode, intBucketId, dblContribution) VALUES " & _
'             "('" & assetname & "','" & bucketId & "'," & dblContribution & ")"
'    Call execCommandSQL(strSQL)
'End Sub

' associa un catSwapLayer con un bucket Katarsis:
' inserisce un record nella tabella (tblAssetBucket)

Public Sub insert_Asset_BucketKatarsis(ByVal assetname As String, ByVal bucketName As String, _
                    ByVal dblContribution As Double)
                    
    Dim strSQL As String

    If checkStringKeyExists("tblBucket", "strName", bucketName) = False Then
        MsgBox "Error: key <" & bucketName & "> not found in table tblBucket (strName)" & _
                vbCrLf & vbCrLf & "Il Bucket <" & bucketName & "> non è presente nella tabella Buckets"
        Exit Sub
    End If
    
    strSQL = "SELECT intid INTO @bucketid FROM tblbucket WHERE strName='" & bucketName & "';"
    Call execCommandSQL(strSQL)
    
    strSQL = "INSERT INTO tblAssetBucket(strAssetCode, intBucketId, dblContribution) VALUES " & _
             "('" & assetname & "',@bucketid," & dblContribution & ")"
    
    Call execCommandSQL(strSQL)
    
End Sub


' rimuove tutti i bucket Hannover associati ad un layer
'Public Sub delete_CatSwapLayer_BucketKatarsis(strLayerName As String, Optional doErrorCheck As Boolean = False)
'    Call deleteStringKey("tblAssetBucket", "strAssetCode", strLayerName, doErrorCheck)
'End Sub

















