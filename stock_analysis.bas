Attribute VB_Name = "Module1"
Sub Stock_analysis()

Dim ws As Worksheet
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim quarterly_change As Double
Dim percent_change As Double
Dim last_row As Long
Dim i As Long
Dim summary_row As Long
Dim max_ticker_increase As String
Dim max_ticker_decrease As String
Dim max_ticker_volume As String
Dim total_volume As Double
Dim max_increase As Double
Dim max_decrease As Double
Dim max_volume As Double

For Each ws In Worksheets
    
    last_row = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    summary_row = 2
    ticker = 0

    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Quarterly Change"
    ws.Cells(1, "K").Value = "Percentage Change"
    ws.Cells(1, "L").Value = "Total Volume"
    ws.Cells(1, "O").Value = "Metric"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"

    For i = 2 To last_row
        If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
            
            ticker = ws.Cells(i, "A").Value
            
            close_price = ws.Cells(i, "F").Value
            
            total_volume = total_volume + ws.Cells(i, "G").Value
            
            quarterly_change = close_price - open_price
            
            percent_change = (quarterly_change / open_price) * 100

            ws.Cells(summary_row, "I").Value = ticker
            ws.Cells(summary_row, "J").Value = quarterly_change
            ws.Cells(summary_row, "K").Value = percent_change
            ws.Cells(summary_row, "L").Value = total_volume

            If quarterly_change > 0 Then
                ws.Cells(summary_row, "J").Interior.ColorIndex = 4
            
            ElseIf quarterly_change < 0 Then
                ws.Cells(summary_row, "J").Interior.ColorIndex = 3
            End If

            If percent_change > max_increase Then
                max_increase = percent_change
                max_ticker_increase = ticker
            End If

            If percent_change < max_decrease Then
                max_decrease = percent_change
                max_ticker_decrease = ticker
            End If

            If total_volume > max_volume Then
                max_volume = total_volume
                max_ticker_volume = ticker
            End If

              
            summary_row = summary_row + 1
        Else
            
            total_volume = total_volume + ws.Cells(i, "G").Value
            
            If ws.Cells(i - 1, "A").Value <> ws.Cells(i, "A").Value Then
                open_price = ws.Cells(i, "C").Value
            End If
        
        End If
    Next i

    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(2, "P").Value = max_ticker_increase
    ws.Cells(2, "Q").Value = max_increase
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(3, "P").Value = max_ticker_decrease
    ws.Cells(3, "Q").Value = max_decrease
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    ws.Cells(4, "P").Value = max_ticker_volume
    ws.Cells(4, "Q").Value = max_volume

Next ws

End Sub

