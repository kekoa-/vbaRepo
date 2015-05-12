Attribute VB_Name = "modUpdateRef"
Option Explicit



' removes the reference to NOAH lib, if exists
Sub RemoveReference(ByRef wb As Workbook)
    Dim vbProj As VBIDE.VBProject
    Dim chkRef As VBIDE.Reference

    Set vbProj = wb.VBProject

    For Each chkRef In vbProj.References
        If chkRef.Name = "NOAH_Lib_1_1" Then
            'On Error Resume Next
            Call vbProj.References.Remove(chkRef)
            'On Error GoTo 0
        End If
    Next

End Sub


' add reference to NOAH LIB
Sub AddReferenceNoah(ByRef wb As Workbook)
     Call Add_Custom_Reference(wb, "K:\Shared Modeling\ANALISI FUNZIONALI\NOAH\NOAH_Lib_1.xlam")
End Sub

' add reference to NOAH LIB DEVELOPMENT
Sub AddReferenceNoahDev(ByRef wb As Workbook)
     Call Add_Custom_Reference(wb, "K:\Shared Modeling\ANALISI FUNZIONALI\NOAH\NOAH_Lib_1_dev.xlam")
End Sub

' add reference to an existing external library
Sub Add_Custom_Reference(ByRef wb As Workbook, ByVal newRefName As String)
    Dim vbProj As VBIDE.VBProject
    Set vbProj = wb.VBProject

    On Error GoTo handleErrAddRef
        vbProj.References.AddFromFile newRefName
    On Error GoTo 0
    Exit Sub
handleErrAddRef:
    MsgBox ("Error adding reference: " & newRefName & " - " & Err.Description)
    On Error GoTo 0
End Sub




' removes NOAH LIB
Public Sub tryRemove()
    Call RemoveReference(ThisWorkbook)
End Sub

' adds NOAH LIB as reference
Public Sub tryAdd()
    Call AddReferenceNoah(ThisWorkbook)
End Sub

' adds NOAH LIB DEV as reference
Public Sub tryAdd_Dev()
    Call AddReferenceNoahDev(ThisWorkbook)
End Sub

' NOAH LIB function call, to check if it s correctly linked
Public Sub tryRef()
    Call logToFile("DoneXX")
End Sub


