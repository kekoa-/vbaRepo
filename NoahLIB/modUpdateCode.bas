Attribute VB_Name = "modUpdateCode"
Option Explicit

' crea le directory, ed esporta i moduli
Sub defaultExport()
    Dim strTime As String
    Dim strPath As String
    On Error Resume Next
        strTime = Format(Now, " yyyy-mm-dd hh.mm.ss")
        MkDir (ThisWorkbook.Path & "\Source")
        strPath = ThisWorkbook.Path & "\Source\out_" & ThisWorkbook.Name & "_" & strTime
        MkDir (strPath)
    On Error GoTo 0
    Call modUpdateCode_export(strPath)
End Sub


'
'
'   MAIN FUNCTION list:
'
'   modUpdateCode_import:    imports the modules into the VB Project, from the file "moduleList.txt", or from user-specified file
'   modUpdateCode_clean:     clean the modules: removes the modules from the VB Project (only removes modules listed in  "moduleList.txt", or user-specified file)
'   modUpdateCode_export:    exports the modules in ".\source_out", or user-specified directory
'
'
'   helper FUNCTION LIST
'
'   LoadSource:  carica i moduli nel VBA Project, leggendo la lista da file
'   CodeClean:  rimuove i moduli
'   CodeExport: esporta i moduli
'
'   ImportCodeModule:   importa un singolo modulo
'   getModuleList:      legge da file la lista dei nomi dei moduli da caricare
'   modUpdateCode_version:   version string
'       checkFileName



' LOADS the modules
' if no parameters are specified, uses the workbook working directory, and "moduleList.txt"
Public Sub modUpdateCode_import(Optional ByVal dirPath As String, Optional ByVal fileName As String)
    ' does a backup first..
    Call defaultExport
    Call LoadSource(ThisWorkbook, dirPath, fileName)
End Sub

' CLEAN the container
' if no parameters are specified, uses the workbook working directory, and "moduleList.txt"
Public Sub modUpdateCode_clean(Optional ByVal dirPath As String, Optional ByVal fileName As String)
    Call CodeClean(ThisWorkbook, dirPath, fileName)
End Sub

' EXPORT the container's Modules
' if no parameters are specified, uses the workbook working directory
Public Sub modUpdateCode_export(Optional ByVal dirPath As String)
    Call CodeExport(ThisWorkbook, dirPath)
End Sub

' VERSION on the container
Public Function modUpdateCode_version() As String
    modUpdateCode_version = "modUpdateCode version 1.0"
End Function




'*****************************************************************
'   Function LoadSource
'
'   Apre moduleList.txt (che si trova nella stessa directory del wb),
'   e importa i moduli elencati
'
'

Public Sub LoadSource(ByRef wb As Workbook, Optional ByVal dirPath As String, Optional ByVal fileName As String)

    Dim fileList() As String
    Dim moduleNameList() As String
    Dim numFilesToImport As Long, i As Long
    
    ReDim fileList(1000)
    ReDim moduleNameList(1000)
    
    If (Len(dirPath) = 0) Then
        dirPath = wb.Path
    End If
    
    If (Len(fileName) = 0) Then
        fileName = "moduleList.txt"
    End If
    
    'get filenames
    Call getModuleList(dirPath & "\" & fileName, moduleNameList, fileList, numFilesToImport)
    
    'cicla sui files
    For i = 0 To numFilesToImport - 1

        If fileList(i) <> "" Then
            ' imports the module
            Call ImportCodeModule(wb, dirPath & "\", moduleNameList(i), fileList(i))
        End If
        
    Next
    
End Sub


'*****************************************************************
'   Function CodeClean
'
'   Apre la lista dei moduli
'   e RIMUOVE i moduli elencati
'

Public Sub CodeClean(ByRef wb As Workbook, Optional ByVal dirPath As String, Optional ByVal fileName As String)

    Dim fileList() As String
    Dim moduleNameList() As String
    Dim numFilesToImport As Integer
    
    ReDim fileList(1000)
    ReDim moduleNameList(1000)
    
    
    If (Len(dirPath) = 0) Then
        dirPath = wb.Path
    End If
    
    If (Len(fileName) = 0) Then
        fileName = "moduleList.txt"
    End If
    
    'get filenames
    Call getModuleList(dirPath & "\" & fileName, moduleNameList, fileList, numFilesToImport)
    
    'cicla sui files
    For i = 0 To numFilesToImport - 1
        If moduleNameList(i) <> "" Then
            ' RIMUOVE the module
            Call RemoveCodeModule(wb, moduleNameList(i))
        End If
        
    Next
    
End Sub


' importa il modulo indicato, nel workbook/VB project
Sub ImportCodeModule(ByRef wb As Workbook, ByVal dirPath As String, ByVal ModuleName As String, ByVal fileName As String)
    ' rimuove il module, se esiste già
    Call RemoveCodeModule(wb, ModuleName)
    ' importa il module
    With wb.VBProject
        If checkFileName(dirPath & fileName & ".bas") Then
            .VBComponents.Import fileName:=dirPath & fileName & ".bas"
        End If
    End With
End Sub


' RIMUOVE il modulo indicato, dal workbook/VB project
Sub RemoveCodeModule(ByRef wb As Workbook, ByVal ModuleName As String)
    With wb.VBProject
        On Error Resume Next
        ' rimuove il module
        .VBComponents.Remove .VBComponents(ModuleName)
        On Error GoTo 0
    End With
End Sub




' apre il file indicato (fileName)
' legge i files e li mette in 'list'
Public Sub getModuleList(ByVal fileName, listModuleName() As String, listFile() As String, ByRef n As Integer)
    Dim s As String
    Dim fs, a
    Dim i As Integer
    
    If (Not (checkFileName(fileName))) Then Exit Sub
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.OpenTextFile(fileName, 1)
    
    i = 0
    Do While a.AtEndofStream <> True
       s = a.readline
       s = Replace(s, vbTab, " ")
       
       listModuleName(i) = Trim(Split(s)(0))
       listFile(i) = Trim(Split(s)(1))
       i = i + 1
    Loop
    a.Close
    
    ' restituisce il numero di files letti
    n = i

End Sub



'exports the modules
Public Sub CodeExport(wb As Workbook, Optional ByVal dirPath As String)
    If (Len(dirPath) = 0) Then
        dirPath = wb.Path & "\Source\out"
    End If
    
    Call exportModules(wb, dirPath)
End Sub



' exports the code modules into the specified path
Sub exportModules(ByRef wb As Workbook, ByVal strPath As String)

Dim i, sName

With wb.VBProject
    For i = 1 To .VBComponents.count
        If .VBComponents(i).CodeModule.CountOfLines > 0 Then
            sName = .VBComponents(i).CodeModule.Name
            .VBComponents(i).Export strPath & "\" & sName & ".bas"
        End If
    Next i
End With

End Sub



'checks the filename
Public Function checkFileName(ByVal fileName As String)
    Dim oFSO
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    If (oFSO.FileExists(fileName)) Then
        checkFileName = True
    Else
        MsgBox ("Error: " & fileName & " not found" & vbCrLf & "in: modUpdateCode")
        checkFileName = False
    End If
End Function



Public Sub defaultImport()
        Call modUpdateCode_import
End Sub






