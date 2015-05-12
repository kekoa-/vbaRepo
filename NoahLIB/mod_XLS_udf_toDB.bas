Attribute VB_Name = "mod_XLS_udf_toDB"


Option Explicit
Option Compare Text


' public getData helper object
Public oGDH As cGetData_helper

' check helper, if not inittialized then initializes it
Public Sub checkHelper()
    If oGDH Is Nothing Then
        Set oGDH = New cGetData_helper
        Call oGDH.init
    End If
End Sub

'
' this is the generic data retrieve function, other function build on this function
Public Function getData( _
    ByVal strcode As String, _
    ByVal strCommand As String, _
    ByVal strKey As String, _
    Optional bucket As String _
)

Dim strSQL As String
'Application.Volatile (True)

getData = "## Invalid Input"
On Error GoTo exitFunction1

'#####################################################################
' use the helper object to retrieve the PLAIN DATA
Call checkHelper
getData = oGDH.getData(strcode, strCommand, strKey)
' end helper
'#####################################################################



'#####################################################################
' additional code to retrieve CUSTOM / non plain data

If strcode = "ASSET" And strCommand = "bucket" Then
    strKey = Replace(strKey, "'", "_")
    bucket = Replace(bucket, "'", "_")
    execCommandSQL ("set @bucketID = -1;")
    execCommandSQL ("select intid into @bucketID from tblbucket where strname ='" & bucket & "';")
    If execScalarSQL("select @bucketID;") = -1 Then
        getData = " error, non-existent bucket name <" & bucket & "> "
        GoTo exitFunction1
    End If
    ' check, if empty return zero
    strSQL = " select count(*) from tblassetbucket where strAssetCode = '" & strKey & "' and " & _
             " intbucketid=@bucketID ;"
    If execScalarSQL(strSQL) = 0 Then
        getData = 0
        GoTo exitFunction1
    End If
    
    strSQL = " select dblContribution from tblassetbucket where strAssetCode = '" & strKey & "' and " & _
             " intbucketid=@bucketID ;"
    getData = execScalarSQL(strSQL)
    GoTo exitFunction1
End If


If strcode = "ASSET" Then
    If strCommand = "FIND" Then
        strKey = Replace(strKey, " ", "%")
        strSQL = " select strCode from tblasset where (strNick like '%" & strKey & "%') or (strName like '%" & strKey & "%') order by strName limit 1;"
        getData = execScalarSQL(strSQL)
    Else
        ' ..
    End If
    
End If


If strcode = "ASSET_RISK" Then
    If strCommand = "COMPUTE_EL_AGG_CT" Then
        getData = execScalarSQL("select funGetAssetEL(funGetAssetId('" & strKey & "'));")
    Else
        '...
    End If
End If


exitFunction1:
    On Error GoTo 0

End Function



' retrieves asset-specific data
Public Function getDataAsset(ByVal strCommand As String, ByVal strKey As String)
    Application.Volatile (True)
    getDataAsset = getData("ASSET", strCommand, strKey)
End Function



' retrieves the price for one asset
Public Function getPrice(ByVal assetCode As String, ByVal datePrice As Date)
    Application.Volatile (True)
    getPrice = execScalarSQL("select funGetPrice(funGetAssetId('" & assetCode & "'), '" & Format(datePrice, "yyyy-mm-dd") & "') ;")
End Function


' retrieves FX rate between two currencies
Public Function getFXRate(ByVal currFrom As String, ByVal currTo As String, ByVal dateRate As Date)
    Application.Volatile (True)
    getFXRate = execScalarSQL("select funGetRateAtDate('" & currFrom & "','" & currTo & "','" & Format(dateRate, "yyyy-mm-dd") & "') ;")
End Function



' retrieves the Cat Bonds anagrafica table
Public Sub getCBAnagraficaTable(ByRef rg As Range, ByRef rg_header As Range)
    Dim strSQL  As String
    strSQL = " select * from tblcatbondinfo order by strName;"
    Call putRsIntoRange_noResize(strSQL, rg, rg_header)
    Call formatRangesByName(rg, rg_header)
End Sub


' retrieves the tblAsset table
Public Sub getAssetTable(ByRef rg As Range, ByRef rg_header As Range)
    Dim strSQL  As String
    strSQL = " select * from tblasset order by strAssetType, strName;"
    Call putRsIntoRange_noResize(strSQL, rg, rg_header)
    Call formatRangesByName(rg, rg_header)
End Sub


' retrieves the asset list, search by name or nick
Public Sub getAssetFind(ByRef rg As Range, ByRef rg_header As Range, ByVal searchString)
    Dim strSQL  As String
    searchString = Replace(searchString, " ", "%")
    strSQL = " select strCode, strNick, strName from tblasset where (strNick like '%" & searchString & "%') or (strName like '%" & searchString & "%') order by strAssetType, strName;"
    Call putRsIntoRange_noResize(strSQL, rg, rg_header)
    Call formatRangesByName(rg, rg_header)
End Sub



' retrieves the asset list, search by name or nick
Public Sub getCommandFind(ByRef rg As Range, ByRef rg_header As Range, ByVal searchString)
    Dim strSQL  As String
    strSQL = " call prcGetHelpCommand('" & searchString & "');"
    Call putRsIntoRange_noResize(strSQL, rg, rg_header)
    Call formatRangesByName(rg, rg_header)
End Sub


' retrieves the asset list, search by name or nick
Public Sub getTradeList(ByRef rg As Range, ByRef rg_header As Range, ByVal searchString)
    Dim strSQL  As String
    DB_Utilities.execCommandSQL (" select 0 into @fundid;")
    DB_Utilities.execCommandSQL (" select intId into @fundid from tblfundlist where strname= '" & searchString & "';")
    strSQL = "select intId from tblMovements where intfund=@fundid;"
    Call putRsIntoRange_noResize(strSQL, rg, rg_header)
    Call formatRangesByName(rg, rg_header)
End Sub





Public Function computeValue( _
    ByVal strcode As String, _
    ByVal strCommand As String, _
    ByVal strKey As String _
)
    computeValue = getData(strcode, strCommand, strKey)
End Function



