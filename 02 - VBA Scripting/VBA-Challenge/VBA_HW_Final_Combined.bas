Sub stockmarketanalyst()

' Base assignment
Dim ws As Worksheet
For Each ws In Worksheets

' Create & print headers for variables
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

' Define variables
    Dim ticker As String
    Dim ticker_total As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim open_price As Double
    Dim closing_price As Double
    Dim total_stock_volume As Double
    Dim summary_table_Row As Long
    Dim original_price As Long
    Dim LastRow As Long
 
' Determine last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Set values for variables
        ticker_total = 0
        total_stock_volume = 0
        summary_table_Row = 2
        original_price = 2
    
For i = 2 To LastRow
' Calculate sum of ticker symbol/letters & display value
    ticker_total = ticker_total + ws.Cells(i, 7).Value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ticker = ws.Cells(i, 1).Value

    ws.Range("I" & summary_table_Row).Value = ticker
    ws.Range("L" & summary_table_Row).Value = ticker_total
    ticker_total = 0

' Calculating totals/values for new columns
    open_price = ws.Range("C" & original_price)
    closing_price = ws.Range("F" & i)
    yearly_change = closing_price - open_price
    ws.Range("J" & summary_table_Row).Value = yearly_change

If open_price = 0 Then
    percent_change = 0

Else
    open_price = ws.Range("C" & original_price)
    percent_change = yearly_change / open_price

End If
' Formatting summary table as %
    ws.Range("K" & summary_table_Row).Value = percent_change
    ws.Range("K" & summary_table_Row).NumberFormat = "0.00%"
   

    ' Conditional formatting
    If ws.Range("J" & summary_table_Row).Value >= 0 Then
        ws.Range("J" & summary_table_Row).Interior.ColorIndex = 4

    Else
        ws.Range("J" & summary_table_Row).Interior.ColorIndex = 3

    End If

    summary_table_Row = summary_table_Row + 1
    original_price = i + 1
    End If

Next i
ws.Columns("A:L").AutoFit
Next ws
End Sub

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