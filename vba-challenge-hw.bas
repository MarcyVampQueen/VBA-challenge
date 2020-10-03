Attribute VB_Name = "Module1"
Sub stonks()
    'Define variables
    Dim currentTicker As String
    Dim nextTicker As String
    Dim startPrice As Double
    Dim totalVolume As Double
    Dim change As Double
    Dim rowCount As Long
    Dim outputRow As Integer
    Dim i As Long
    Dim greatInc As Double
    Dim greatDec As Double
    Dim greatVol As Double
    Dim wsName As String
    'had trouble with zeros...this could also be called percentChange
    Dim zeroSux As Double
    
    For Each ws In Worksheets
        'Initialize variables
        rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
        outputRow = 2
        startPrice = ws.Cells(2, 3).Value
        totalVolume = ws.Cells(2, 7).Value
        change = 0
        greatInc = 0
        greatDec = 0
        greatVol = 0
        zeroSux = 0
        ' label the columns on the sheet
        ws.Range("I1:P1").Value = Array("Ticker Symbol", "Yearly Change", "Percent Change", "Total Stock Volume", "", "", "Ticker", "Volume")
        ws.Range("N2:N4").Value = Array("Greatest % increase", "Greatest % decrease", "Greatest total volume")
        
        'Loop through the rows
        For i = 2 To rowCount
            'Get the current row, next row, and add to the total
            currentTicker = ws.Cells(i, 1).Value
            nextTicker = ws.Cells(i + 1, 1).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            'If we're changing tickers, output the summarized information and format it
            If currentTicker <> nextTicker Then
                change = ws.Cells(i, 6).Value - startPrice
                'Output the info
                ws.Cells(outputRow, 9).Value = currentTicker
                ws.Cells(outputRow, 10).Value = change
                If startPrice <> 0 Then
                    zeroSux = change / startPrice
                Else
                    'not sure this is actually the right thing to do if the %change was effectively infinite cuz started at 0
                    zeroSux = 0
                End If
                ws.Cells(outputRow, 11).Value = zeroSux
                ws.Cells(outputRow, 12).Value = totalVolume
                
                'Is it the greatest volume we've seen yet?
                If totalVolume > greatVol Then
                    greatVol = totalVolume
                    Range("O4").Value = currentTicker
                    Range("P4").Value = greatVol
                End If
                
                'Format the row and look for the greatest increase/decrease
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                If change >= 0 Then
                    'green; is it the greatest increase?
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4
                    If (zeroSux) > greatInc Then
                        greatInc = zeroSux
                        Range("O2").Value = currentTicker
                        Range("P2").Value = greatInc
                        Range("P2").NumberFormat = "0.00%"
                    End If
                Else
                    'red; is it the greatest decrease?
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3
                    If (zeroSux) < greatDec Then
                        greatDec = zeroSux
                        Range("O3") = currentTicker
                        Range("P3").Value = greatDec
                        Range("P3").NumberFormat = "0.00%"
                    End If
                End If
                
                'Make sure we don't overwrite data, reset the volume, and set the new starting price
                outputRow = outputRow + 1
                totalVolume = 0
                startPrice = ws.Cells(i + 1, 3)
            End If
        Next i
    Next
    
End Sub
