Attribute VB_Name = "mod_insertCells"
Option Explicit

' sposta la tabella, se esiste
Public Sub tryMoveText6(ByRef ws As Worksheet)
    Dim shape
    On Error Resume Next
    Set shape = ActiveSheet.Shapes("TextBox 6")
    shape.Left = shape.Left + 60
    On Error GoTo 0
End Sub

'inserisce una colonna nella tabella "RISK MEASURES"
Public Sub updateRange_ELAggregateAIR(ByRef ws As Worksheet)
    Dim rg As Range, rg1 As Range, rg2 As Range, rg3 As Range, rg4 As Range
    
    
    On Error Resume Next
    ' check if range exists
    Set rg = ActiveSheet.Range("rng_AIR_RiskMeasure_ELaggregate")
    On Error GoTo 0
    'if range exist, exit
    If Not rg Is Nothing Then
        Exit Sub
    End If
    
    
    ' if range does not exist, insert it
    Set rg1 = Range("rng_AIR_RiskAmount")
    Set rg2 = ws.Cells.Range(rg1.Cells(-1, 1), rg1.Cells(rg1.Rows.count, rg1.Columns.count))
    Set rg2 = rg2.Columns(1)
        rg2.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Set rg3 = rg2.Cells(2, 0)
        Set rg4 = ws.Cells.Range(rg3, rg3.Cells(rg2.Rows.count - 2, 1))
        Call formatRange_outSideBorders(rg4)
        
    Set rg3 = ws.Cells.Range(rg2.Cells(3, 0), rg2.Cells(rg2.Rows.count, 0))
    Call formatRange_topBorders(rg3)
    
    ActiveWorkbook.Names.add Name:="rng_AIR_RiskMeasure_ELaggregate", _
       RefersToR1C1:="=AIR!" & rg3.Address(ReferenceStyle:=xlR1C1)
     
    rg3.Cells(0, 1).value = "EL Aggregate"
    
    Call tryMoveText6(ws)

End Sub



'inserisce una colonna nella tabella "RISK MEASURES"
Public Sub updateRange_ExhProbAggregateAIR(ByRef ws As Worksheet)
    Dim rg As Range, rg1 As Range, rg2 As Range, rg3 As Range, rg4 As Range
    
    
    On Error Resume Next
    ' check if range exists
    Set rg = ActiveSheet.Range("rng_AIR_RiskMeasure_ExhProbaggregate")
    On Error GoTo 0
    'if range exist, exit
    If Not rg Is Nothing Then
        Exit Sub
    End If
    
    
    ' if range does not exist, insert it
    Set rg1 = Range("rng_AIR_RiskAmount")
    Set rg2 = ws.Cells.Range(rg1.Cells(-1, 1), rg1.Cells(rg1.Rows.count, rg1.Columns.count))
    Set rg2 = rg2.Columns(1)
        rg2.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Set rg3 = rg2.Cells(2, 0)
        Set rg4 = ws.Cells.Range(rg3, rg3.Cells(rg2.Rows.count - 2, 1))
        Call formatRange_outSideBorders(rg4)
        
    Set rg3 = ws.Cells.Range(rg2.Cells(3, 0), rg2.Cells(rg2.Rows.count, 0))
    Call formatRange_topBorders(rg3)
    
    ActiveWorkbook.Names.add Name:="rng_AIR_RiskMeasure_ExhProbaggregate", _
       RefersToR1C1:="=AIR!" & rg3.Address(ReferenceStyle:=xlR1C1)
     
    rg3.Cells(0, 1).value = "Exh Prob Aggregate"
    rg3.Cells(0, 1).WrapText = True
    
    Call tryMoveText6(ws)

End Sub



'inserisce una colonna nella tabella "RISK MEASURES"
Public Sub updateRange_AttProbAggregateAIR(ByRef ws As Worksheet)
    Dim rg As Range, rg1 As Range, rg2 As Range, rg3 As Range, rg4 As Range
    
    
    On Error Resume Next
    ' check if range exists
    Set rg = ActiveSheet.Range("rng_AIR_RiskMeasure_AttProbaggregate")
    On Error GoTo 0
    'if range exist, exit
    If Not rg Is Nothing Then
        Exit Sub
    End If
    
    
    ' if range does not exist, insert it
    Set rg1 = Range("rng_AIR_RiskMeasure_ELaggregate")
    Set rg2 = ws.Cells.Range(rg1.Cells(-1, 1), rg1.Cells(rg1.Rows.count, rg1.Columns.count))
    Set rg2 = rg2.Columns(1)
        rg2.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Set rg3 = rg2.Cells(2, 0)
        Set rg4 = ws.Cells.Range(rg3, rg3.Cells(rg2.Rows.count - 2, 1))
        Call formatRange_outSideBorders(rg4)
        
    Set rg3 = ws.Cells.Range(rg2.Cells(3, 0), rg2.Cells(rg2.Rows.count, 0))
    Call formatRange_topBorders(rg3)
    
    ActiveWorkbook.Names.add Name:="rng_AIR_RiskMeasure_AttProbaggregate", _
       RefersToR1C1:="=AIR!" & rg3.Address(ReferenceStyle:=xlR1C1)
     
    rg3.Cells(0, 1).value = "Att Prob Aggregate"
    rg3.Cells(0, 1).WrapText = True
    
    Call tryMoveText6(ws)

End Sub


Public Sub updateRange_Format_AggregateAIR(ByRef ws As Worksheet)
    Dim rg1 As Range, rg2 As Range, rg3 As Range, rg4 As Range
    
    Set rg1 = ActiveSheet.Range("rng_AIR_RiskMeasure_AttProbaggregate").Cells(0, 1)
    Set rg2 = ActiveSheet.Range("rng_AIR_RiskMeasure_ExhProbaggregate").Cells(0, 1)
    Set rg3 = ws.Cells.Range(rg1, rg2)
    'rg3.Select
    Call formatRange_outSideBorders(rg3)
    
    Set rg1 = ActiveSheet.Range("rng_AIR_RiskMeasure_AttProbaggregate").Cells(1, 1)
    Set rg2 = ActiveSheet.Range("rng_AIR_RiskMeasure_ExhProbaggregate")
    Set rg3 = rg2.Cells(rg2.Rows.count - 1, 1)
    Set rg4 = ws.Cells.Range(rg1, rg3)
    'rg4.Select
    Call formatRange_outSideBorders(rg4)
    
    
End Sub




'inserisce una colonna nella tabella "RISK MEASURES"
Public Sub updateRange_ELAggregateRMS(ByRef ws As Worksheet)
    Dim rg As Range, rg1 As Range, rg2 As Range, rg3 As Range, rg4 As Range
    
    
    
    On Error Resume Next
    ' check if range exists
    Set rg = ActiveSheet.Range("rng_RMS_RiskMeasure_ELaggregate")
    On Error GoTo 0
    'if range exist, exit
    If Not rg Is Nothing Then
        Exit Sub
    End If
    
    
    ' if range does not exist, insert it
    Set rg1 = Range("rng_RMS_RiskAmount")
    Set rg2 = ws.Cells.Range(rg1.Cells(-1, 1), rg1.Cells(rg1.Rows.count, rg1.Columns.count))
    Set rg2 = rg2.Columns(1)
        rg2.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Set rg3 = rg2.Cells(2, 0)
        Set rg4 = ws.Cells.Range(rg3, rg3.Cells(rg2.Rows.count - 2, 1))
        Call formatRange_outSideBorders(rg4)
        
    Set rg3 = ws.Cells.Range(rg2.Cells(3, 0), rg2.Cells(rg2.Rows.count, 0))
    
    Call formatRange_topBorders(rg3)
    
    ActiveWorkbook.Names.add Name:="rng_RMS_RiskMeasure_ELaggregate", _
       RefersToR1C1:="=RMS!" & rg3.Address(ReferenceStyle:=xlR1C1)
     
    rg3.Cells(0, 1).value = "EL Aggregate"

End Sub



Public Sub formatRange_outSideBorders(ByRef rg As Range)
        With rg.Borders
            .LineStyle = xlNone
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        rg.Borders(xlInsideVertical).LineStyle = xlNone
        rg.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

Public Sub formatRange_topBorders(ByRef rg As Range)
    rg.Borders(xlEdgeTop).LineStyle = xlNone
    rg.Borders(xlEdgeTop).ColorIndex = 0
    rg.Borders(xlEdgeTop).TintAndShade = 0
    rg.Borders(xlEdgeTop).Weight = xlThin
End Sub





