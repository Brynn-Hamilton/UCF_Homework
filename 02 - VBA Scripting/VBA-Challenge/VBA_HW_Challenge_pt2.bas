Sub VBA_Stocks_Challenge()

'Challenge
Dim ws As Worksheet
For Each ws In Worksheets


' Create & print headers for summary table
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker Total"
    ws.Range("Q1").Value = "Value"

' Define summary table variables
    Dim greatest_percent_increase As Double
    Dim greatest_percent_decrease As Double
    Dim greatest_total_volume As Double
    Dim summary_table_Row As Long

    'Set values for summary table variables
        greatest_percent_increase = 0
        greatest_percent_decrease = 0
        greatest_total_volume = 0
        summary_table_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Final format for summary table as %
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"

    For i = 2 To summary_table_Row
' Final calculations for greatest % increase/decrease/total
If greatest_total_volume <= ws.Range("L" & i).Value Then
    greatest_total_volume = ws.Range("L" & i).Value
    ws.Range("P4").Value = ws.Range("I" & i).Value
    ws.Range("Q4").Value = ws.Range("L" & i).Value
    ElseIf greatest_percent_increase <= ws.Range("K" & i).Value Then
      greatest_percent_increase = ws.Range("K" & i).Value
      ws.Range("P2").Value = ws.Range("I" & i).Value
      ws.Range("Q2").Value = ws.Range("K" & i).Value
    ElseIf greatest_percent_decrease >= ws.Range("K" & i).Value Then
      greatest_percent_decrease = ws.Range("K" & i).Value
      ws.Range("P3").Value = ws.Range("I" & i).Value
      ws.Range("Q3").Value = ws.Range("K" & i).Value
    End If
 Next i
 ws.Range("Q4").NumberFormat = "0.00E+00"
 ws.Columns("O:Q").AutoFit
    
Next ws

End Sub