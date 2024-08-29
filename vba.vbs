Sub ticker()
    
    Dim i As Long
    Dim j As Long
    Dim ticker As String
    Dim lastrow As Long
    Dim ws As Worksheet
    Dim tickerrow As Integer
    Dim qc As Double
    Dim qcrow As Integer
    Dim price_change As Double
    Dim pcrow As Integer
    Dim ts_vol As Double
    Dim tsrow As Integer
    Dim great_inc As Double
    Dim great_dec As Double
    Dim great_tv As Double
    Dim last_tick_row As Long
    
    For Each ws In ThisWorkbook.Worksheets
        
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ticker = " "
        tickerrow = 2
        j = 2
        qc = 0
        qcrow = 2
        price_change = 0
        pcrow = 2
        ts_vol = 0
        tsrow = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
      For i = 2 To lastrow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
             ' Set the ticker
             ticker = ws.Cells(i, 1).Value

             ' Print ticker in column I
             ws.Range("I" & tickerrow).Value = ticker
             
             ' Calculate quarterly change
             qc = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

             ' Print the quarterly change in column J
             ws.Range("J" & qcrow).Value = qc
             
             ' color the quarterly change calculations
                If qc < 0 Then
                        ws.Range("J" & qcrow).Interior.ColorIndex = 3
                    ElseIf qc > 0 Then
                        ws.Range("J" & qcrow).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & qcrow).Interior.ColorIndex = 0
                End If
    
             ' Calculate percent change
             percent_change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
             
             ' print percent change in column k and format as percent
             ws.Range("K" & pcrow).Value = percent_change
             ws.Range("K" & pcrow).NumberFormat = "0.00%"
             
             ' Calculate total stock volume
             ts_vol = ws.Application.WorksheetFunction.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(i, 7)))
             
             ' print total stock volume in column m
             ws.Range("L" & tsrow).Value = ts_vol
             
             ' Add one to the row
             tickerrow = tickerrow + 1
             qcrow = qcrow + 1
             pcrow = pcrow + 1
             tsrow = tsrow + 1

             ' Reset variables
             qc = 0
             j = i + 1
                 
            
        End If
        
      Next i
              
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        last_tick_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        
        great_inc = ws.Cells(2, 11).Value
        great_dec = ws.Cells(2, 11).Value
        great_tv = ws.Cells(2, 12).Value
        
        For i = 3 To last_tick_row
        
        ' find greatest percent increase
        If ws.Cells(i, 11) > great_inc Then
            great_inc = ws.Cells(i, 11).Value
            ws.Range("Q2").Value = great_inc
            ws.Range("P2").Value = ws.Cells(i, 9).Value
        End If
            
        ' find greatest percent decrease
        If ws.Cells(i, 11) < great_dec Then
            great_dec = ws.Cells(i, 11).Value
            ws.Range("Q3").Value = great_dec
            ws.Range("P3").Value = ws.Cells(i, 9).Value
            ws.Range("Q3").NumberFormat = "0.00%"
        End If
        
        ' find greatest total volume
        If ws.Cells(i, 12) > great_tv Then
            great_tv = ws.Cells(i, 12).Value
            ws.Range("Q4").Value = great_tv
            ws.Range("P4").Value = ws.Cells(i, 9).Value
        End If
             
        Next i
        
    Next ws
    
End Sub

