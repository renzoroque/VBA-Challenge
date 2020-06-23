Attribute VB_Name = "Module1"
Sub Calculate_All()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call ticker_calc
    Next
    Application.ScreenUpdating = True
End Sub

Sub ticker_calc():

Dim ticker As String

Dim open_price As Double
open_price = Range("C2").Value

Dim close_price As Double

Dim price_change As Double

Dim percent_change As Double

Dim total_volume As Double
total_volume = 0

Dim table_row As Integer
table_row = 2

last_row = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To last_row
        If Cells(i + 1, 1) <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        close_price = Cells(i, 6).Value
        
        price_change = close_price - open_price
            If open_price = 0 Then
            percent_change = 0
            Else
            percent_change = (close_price - open_price) / open_price
            End If
            
        total_volume = total_volume + Cells(i, 7).Value

        Range("J" & table_row).Value = ticker
        Range("K" & table_row).Value = price_change
            If price_change < 0 Then
            Cells(table_row, 11).Interior.ColorIndex = 3
            Else
            Cells(table_row, 11).Interior.ColorIndex = 4
            End If
            
        Range("L" & table_row).Value = percent_change
        Cells(table_row, 12).NumberFormat = "0.00%"
        Range("M" & table_row).Value = total_volume
        
        table_row = table_row + 1
        
        open_price = Cells(i + 1, 3).Value
        
        total_volume = 0
        
        Else
        
        total_volume = total_volume + Cells(i, 7).Value
                
        End If
        
    Next i
    
Dim max_perc As Double
max_perc = Range("L2").Value

Dim min_perc As Double
min_perc = Range("L2").Value

Dim max_ticker_sym As String

Dim min_ticker_sym As String

Dim max_volume As Double
max_volume = Range("m2").Value

Dim volume_ticker As String

    For h = 2 To last_row
    
        If max_perc < Cells(h + 1, 12) Then
        max_perc = Cells(h + 1, 12).Value
        max_ticker_sym = Cells(h + 1, 10).Value
        
        End If
        
           
        If min_perc > Cells(h + 1, 12) Then
        min_perc = Cells(h + 1, 12).Value
        min_ticker_sym = Cells(h + 1, 10).Value
               
        End If
        
        If max_volume < Cells(h + 1, 13) Then
        max_volume = Cells(h + 1, 13).Value
        volume_ticker = Cells(h + 1, 10).Value
        
        End If
                                
    Next h
    
    Range("Q4").Value = max_volume
    Range("Q2").Value = max_perc
    Cells(2, 17).NumberFormat = "0.00%"
    Range("Q3").Value = min_perc
    Cells(3, 17).NumberFormat = "0.00%"
    Range("P2").Value = max_ticker_sym
    Range("P3").Value = min_ticker_sym
    Range("P4").Value = volume_ticker
    
    
    
'Headers

Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"


End Sub


