Attribute VB_Name = "Module1"
Sub StockFlow()

Dim ws As Worksheet
Dim ticker_name As String
Dim percent_change As Double
Dim total_volume As Double
Dim year_open As Double
Dim year_close As Double
Dim lastrow As Long
Dim year_change As Double

For Each ws In Worksheets


ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

total_volume = 0
ticker_column = 2

For i = 2 To lastrow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    
    
        ticker_name = ws.Cells(i, 1).Value
        total_volume = total_volume + ws.Cells(i, 7).Value
        ws.Cells(ticker_column, 9).Value = ticker_name
        
        year_open = ws.Cells(i - 252, 3).Value
        year_close = ws.Cells(i, 6).Value
        year_change = year_close - year_open
        percent_change = year_change / year_open
    
        ws.Cells(ticker_column, 10).Value = year_change
        
    If ws.Cells(ticker_column, 10).Value < 0 Then
            ws.Cells(ticker_column, 10).Interior.Color = vbRed
    
    Else
            ws.Cells(ticker_column, 10).Interior.Color = vbGreen
    
    
    End If
    
        ws.Cells(ticker_column, 11).Value = percent_change
        ws.Cells(ticker_column, 11).NumberFormat = "0.00%"
        ws.Cells(ticker_column, 12).Value = total_volume
    
        
 ticker_column = ticker_column + 1
 total_volume = 0
 
 
 Else
    total_volume = total_volume + ws.Cells(i, 7).Value
    
    End If
    
    
    Dim percent_range As Double
    Dim totalstock_volume As Double
    
    
    highestpercent = ws.Application.WorksheetFunction.Max(ws.Range("K:K"))
    lowestpercent = ws.Application.WorksheetFunction.Min(ws.Range("K:K"))
    highestvolume = ws.Application.WorksheetFunction.Max(ws.Range("L:L"))
    
    ws.Range("Q2").Value = highestpercent
    ws.Range("Q3").Value = lowestpercent
    ws.Range("Q4").Value = highestvolume
   
   ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "0"
    
Next i

ws.Range("I:L").EntireColumn.AutoFit



Next ws

End Sub
