Sub StockData():
 
 Dim ws As Worksheet
    Dim i As Long
    Dim rowcount As Long
    Dim uniqueTickers As Collection
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentchange As Double
    Dim openPrice As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim summaryWs As Worksheet
    Dim lastRow As Long
    Dim summaryrow As Double
    

    For Each ws In ThisWorkbook.Sheets
     
            rowcount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            greatestIncrease = 0
            greatestDecrease = 0
            greatestVolume = 0
            greatestIncreaseTicker = ""
            greatestDecreaseTicker = ""
            greatestVolumeTicker = ""
            percentchange = 0
            summaryrow = 2
            
             ' Set the header "Ticker" in cell I1
            ws.Range("I1").Value = "Ticker"
            ' Set the header "Yearly Change" in cell J1
            ws.Range("J1").Value = "Yearly Change"
            ' Set the header "Percent Change" in cell K1
            ws.Range("K1").Value = "Percent Change"
            ' Set the header "Total Stock Volume" in cell L1
            ws.Range("L1").Value = "Total Stock Volume"
            openPrice = ws.Cells(2, 3).Value
            ' Loop through tickers and add unique ones to the collection
            For i = 2 To rowcount
         
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Calculate yearly change
                yearlyChange = ws.Cells(i, 6).Value - openPrice ' Assuming close - open
                ' Calculate open price
                totalVolume = totalVolume + ws.Cells(i, 7).Value
               ' totalVolume = Application.WorksheetFunction.SumIf(ws.Range("A:A"), ws.Cells(i, 1).Value, ws.Range("G:G"))
                 ws.Range("I" & summaryrow).Value = ws.Cells(i, 1).Value
                ' Add the values to the corresponding cells
                ws.Range("J" & summaryrow).Value = yearlyChange
                If openPrice = 0 Then
                percentchange = 0
                Else
                percentchange = yearlyChange / openPrice
                End If
                ws.Range("K" & summaryrow).Value = percentchange
                ws.Range("K" & summaryrow).NumberFormat = "0.00%"
                ws.Range("L" & summaryrow).Value = totalVolume
                
                ' Format color in column J based on yearly change
                If yearlyChange < 0 Then
                    ws.Range("J" & summaryrow).Interior.Color = RGB(255, 0, 0) ' Red
                ElseIf yearlyChange > 0 Then
                    ws.Range("J" & summaryrow).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Range("J" & summaryrow).Interior.Color = RGB(255, 255, 255) ' White
                End If
                ' Find greatest increase, decrease, and volume
                If ws.Range("K" & summaryrow).Value > greatestIncrease Then
                    greatestIncrease = ws.Range("K" & summaryrow).Value
                    greatestIncreaseTicker = ws.Range("I" & summaryrow).Value
                End If
                If ws.Range("K" & summaryrow).Value < greatestDecrease Then
                    greatestDecrease = ws.Range("K" & summaryrow).Value
                    greatestDecreaseTicker = ws.Range("I" & summaryrow).Value
                End If
                If ws.Range("L" & summaryrow).Value > greatestVolume Then
                    greatestVolume = ws.Range("L" & summaryrow).Value
                    greatestVolumeTicker = ws.Range("I" & summaryrow).Value
                End If
                openPrice = ws.Cells(i + 1, 3).Value ' Opening price
                totalVolume = 0
                summaryrow = summaryrow + 1
                
                Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                End If
                
            Next i
            ' Output the greatest increase, decrease, and volume
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("P2").Value = greatestIncreaseTicker
            ws.Range("P3").Value = greatestDecreaseTicker
            ws.Range("P4").Value = greatestVolumeTicker
            ws.Range("Q2").Value = greatestIncrease
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").Value = greatestDecrease
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("Q4").Value = greatestVolume
   
    Next ws
End Sub
