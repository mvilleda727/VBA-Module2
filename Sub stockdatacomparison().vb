Sub stockdatacomparison()
    ' Variables
    Dim ws As Worksheet
    Dim i As Long
    Dim summary_table_row As Long
    Dim ticker_name As String
    Dim Lastrow As Long
    Dim open_value As Double, close_value As Double, totalstock As Double
    Dim change As Double

    ' Loop through worksheets
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        summary_table_row = 2

        ' Column headers
        With ws
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Quarterly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            .Range("Q1").Value = "Value"
            .Range("P1").Value = "Ticker"
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
        End With

        ' Finding last row
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        totalstock = 0

        ' Loop through each row to find unique tickers and calculate changes
        open_value = ws.Cells(2, 3).Value
        For i = 2 To Lastrow
            ' Check if the current ticker is different from the next one
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker_name = ws.Cells(i, 1).Value
                ws.Range("I" & summary_table_row).Value = ticker_name
                
                ' Set values for close, open and total stock
                close_value = ws.Cells(i, 6).Value
                'open_value = ws.Cells(i, 3).Value
            ' Accumulate total stock volume
                totalstock = totalstock + ws.Cells(i, 7).Value 
                
                ' Calculate changes
                 If IsNumeric(close_value) And IsNumeric(open_value) And open_value <> 0 Then
                    change = close_value - open_value
                    ws.Range("J" & summary_table_row).Value = change
                    'percent change
                    ws.Range("K" & summary_table_row).Value = change / open_value
                    
                    ws.Range("L" & summary_table_row).Value = totalstock

                ' Color the change based on value
                    If change > 0 Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4 ' Green
                    ElseIf change < 0 Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3 ' Red
                    End If
                Else
                    ws.Range("J" & summary_table_row).Value = "N/A"
                    ws.Range("K" & summary_table_row).Value = "N/A"
                End If
                
                ' Move to the next row in the summary table
                summary_table_row = summary_table_row + 1
                totalstock = 0 ' Reset for next ticker
                open_value = Cells(i + 1, 3).Value
            Else
                ' Accumulate total stock volume for the current ticker
                totalstock = totalstock + ws.Cells(i, 7).Value
            End If
        Next i
        'Change number format to percentages for entire column
        Range("K2").EntireColumn.NumberFormat = "0.00%"
        ' Last row of the summary table
        Dim summary_lastrow As Long
        summary_lastrow = summary_table_row - 1

        ' Variables to store maximum and minimum
        Dim maxPercentChange As Double
        Dim minPercentChange As Double
        Dim maxVolume As Double
        
        ' Find max and min values
        maxPercentChange = Application.WorksheetFunction.Max(ws.Range("K2:K" & summary_lastrow))
        minPercentChange = Application.WorksheetFunction.Min(ws.Range("K2:K" & summary_lastrow))
        maxVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & summary_lastrow))

        ' Loop through summary table to find the tickers with greatest % increase, decrease, and volume
        Dim k As Long
        For k = 2 To summary_lastrow
            ' Check for greatest % increase
            If ws.Cells(k, 11).Value = maxPercentChange Then
                ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            ' Check for greatest % decrease
            ElseIf ws.Cells(k, 11).Value = minPercentChange Then
                ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            ' Check for greatest total volume
            ElseIf ws.Cells(k, 12).Value = maxVolume Then
                ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
                ws.Cells(4, 17).NumberFormat = "#"
            End If
        Next k
        
        ' Activate Q1 worksheet at the end
        Worksheets("Q1").Activate
    Next ws
End Sub

