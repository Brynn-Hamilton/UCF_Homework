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