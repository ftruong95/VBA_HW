Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data_analysis():

' loop through worksheets

    For Each ws In Worksheets

' label headers and data

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
' set variables

        Dim Ticker As String
        Dim TotalTickerVolume As Double
        TotalTickerVolume = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PreviousAmount As Long
        PreviousAmount = 2
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim LastRow As Long
        Dim LastRowValue As Long
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0

' define last row

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
 
' add/check ticker total volume

            TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryTableRow).Value = Ticker
                ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
                TotalTickerVolume = 0

' calculate yearly change

                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SummaryTableRow).Value = YearlyChange

                If YearlyOpen = 0 Then
                    PercentChange = 0
                
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                
                
                End If

' formating/conditional formatting

                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                 
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                   ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                
                End If

' add to summary table

                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                
                End If
                
            Next i

' loop through for final results

            For i = 2 To LastRow
                
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                
                End If

            Next i
    
    Next ws

End Sub
