Attribute VB_Name = "Module1"
Sub StockData()
    'Loop
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double

    For Each ws In ThisWorkbook.Worksheets
        ' Set initial values
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        TotalVolume = 0
        OpeningPrice = ws.Cells(2, 3).Value

        ' Set headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        Dim intialRow As Integer
        intialRow = 2

        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                ClosingPrice = ws.Cells(i, 6).Value
                
                YearlyChange = ClosingPrice - OpeningPrice
                
                If OpeningPrice <> 0 Then
                    
                    PercentChange = (YearlyChange / OpeningPrice)
                Else
                    PercentChange = 0
                End If
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value

                ' Print to table
                ws.Cells(intialRow, 9).Value = Ticker
                ws.Cells(intialRow, 10).Value = YearlyChange
                ws.Cells(intialRow, 11).Value = PercentChange
                ws.Cells(intialRow, 11).NumberFormat = "0.00%"
                ws.Cells(intialRow, 12).Value = TotalVolume

                ' Conditional formatting
                If YearlyChange > 0 Then
                    ws.Cells(intialRow, 10).Interior.Color = vbGreen
                ElseIf YearlyChange < 0 Then
                    ws.Cells(intialRow, 10).Interior.Color = vbRed
                End If

                intialRow = intialRow + 1
                
                OpeningPrice = ws.Cells(i + 1, 3).Value
                TotalVolume = 0
                Else
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
              ' find the greatest % increase, % decrease, and total volume
        
         
            If YearlyChange > GreatestIncrease Then
                GreatestIncrease = YearlyChange
                IncreaseTicker = Ticker
            ElseIf YearlyChange < GreatestDecrease Then
                GreatestDecrease = YearlyChange
                DecreaseTicker = Ticker
            End If

            If TotalVolume > GreatestVolume Then
                GreatestVolume = TotalVolume
                VolumeTicker = Ticker
            End If

        
        'Output
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = IncreaseTicker
        ws.Cells(3, 16).Value = DecreaseTicker
        ws.Cells(4, 16).Value = VolumeTicker
        ws.Cells(2, 17).Value = GreatestIncrease & "%"
        ws.Cells(3, 17).Value = GreatestDecrease & "%"
        ws.Cells(4, 17).Value = GreatestVolume




            End If
        Next i
    Next ws

End Sub

