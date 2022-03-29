# VBA-challenge
VBA homework - Stefanie Gagnon


Option Explicit

Sub multiyrstockdata()
    Const TICKER_COL As Integer = 1
    Const VOLUME_COL As Integer = 7
    Const OPEN_COL As Integer = 3
    Const CLOSE_COL As Integer = 6
    Const PERCENT_COL As Integer = 11
    Const YRCHNG_COL As Integer = 10
    
    Dim ws As Worksheet
    Dim ticker As String
    Dim stockvolume As Double
    Dim lastrow As Long
    Dim input_row As Long
    Dim output_row As Integer
    Dim opening_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    For Each ws In Worksheets
        ws.Activate
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Stock Volume"
        Columns("J").AutoFit
        Columns("K").AutoFit
        Columns("L").AutoFit
     
        output_row = 2
        For input_row = 2 To lastrow
            If Cells(input_row - 1, TICKER_COL).Value <> Cells(input_row, TICKER_COL).Value Then
                'First row of new ticker
                stockvolume = 0
                opening_price = Cells(input_row, OPEN_COL).Value
            End If
            
            'For every row including first and last
            stockvolume = stockvolume + Cells(input_row, VOLUME_COL).Value
            
            If Cells(input_row + 1, TICKER_COL).Value <> Cells(input_row, TICKER_COL).Value Then
                'Last row of ticker
                ticker = Cells(input_row, TICKER_COL).Value
                Range("I" & output_row).Value = ticker
                Range("L" & output_row).Value = stockvolume
                closing_price = Cells(input_row, CLOSE_COL).Value
                yearly_change = closing_price - opening_price
                Range("J" & output_row).Value = yearly_change
                
                    If Cells(output_row, YRCHNG_COL).Value > 0 Then
                        Cells(output_row, YRCHNG_COL).Interior.ColorIndex = 4
                        End If
                    If Cells(output_row, YRCHNG_COL).Value < 0 Then
                        Cells(output_row, YRCHNG_COL).Interior.ColorIndex = 3
                        End If
                
                percent_change = (yearly_change / opening_price)
                Range("K" & output_row).Value = percent_change
                
                output_row = output_row + 1    'Set up for next stock, must be last!!
            End If
            
            Cells(output_row, PERCENT_COL).NumberFormat = "0.00%"

        Next input_row
       
        
    Next ws
End Sub


