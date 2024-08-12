'Create a script that loops through all the stocks For each quarter And outputs the following information:

'The ticker symbol

'Quarterly change from the opening price at the beginning of a given quarter To the closing price at the end of that quarter.

'The percentage change from the opening price at the beginning of a given quarter To the closing price at the end of that quarter.

Sub MultipleYearStockData()

    Dim SheetName As Variant
    Dim SheetNames As Variant

    SheetNames = Array("Q1", "Q2", "Q3", "Q4")

    For Each SheetName In SheetNames

        Dim ws As Worksheet
        Dim LastRow As Long

        Set ws = ThisWorkbook.Sheets(SheetName)
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = "Greatest Total Volume"


        Dim InputRow As Long
        Dim OutputRow As Long
        Dim Ticker As String
        Dim TargetRow As Long

        TargetRow = 1


        Dim Qchange As Double
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim PercentChange As Double
        'Dim RoundedPercentage As Double
        Dim TotalVolume As Double
        Dim Volume As Double

        For InputRow = 2 To LastRow


            If ws.Cells(InputRow, 1).Value <> Ticker Then
                ' Ticker changes from InputRow - 1 To InputRow

                ' "close" the previous Ticker information

                'Skips first row
                If InputRow > 2 Then
                    ClosePrice = ws.Cells(InputRow - 1, 6).Value
                    Qchange = ClosePrice - OpenPrice
                    ws.Cells(TargetRow, 11).Value = Qchange
                    'Format color
                    If Qchange < 0 Then
                        ws.Cells(TargetRow, 11).Interior.Color = vbRed
                    Elseif Qchange > 0 Then
                        ws.Cells(TargetRow, 11).Interior.Color = vbGreen
                    End If

                    ' Calculate percent change
                    PercentChange = Qchange / OpenPrice
                    ws.Cells(TargetRow, 12).Value = PercentChange
                    ws.Cells(TargetRow, 12).NumberFormat = "0.00%"
                    ws.Cells(TargetRow, 13).Value = TotalVolume



                End If

                ' Then, "open" the Next Ticker information

                Ticker = ws.Cells(InputRow, 1)

                'Get opening price
                TargetRow = TargetRow + 1
                ws.Cells(TargetRow, 10).Value = Ticker
                OpenPrice = ws.Cells(InputRow, "C").Value


                TotalVolume = 0

            End If

            Volume = ws.Cells(InputRow, 7).Value
            TotalVolume = Volume + TotalVolume

        Next InputRow

        'Calculates last row of Quarterly Change And Percent Change
        ClosePrice = ws.Cells(InputRow - 1, 6).Value
        Qchange = ClosePrice - OpenPrice
        ws.Cells(TargetRow, 11).Value = Qchange
        'Format color
        If Qchange < 0 Then
            ws.Cells(TargetRow, 11).Interior.Color = vbRed
        Elseif Qchange > 0 Then
            ws.Cells(TargetRow, 11).Interior.Color = vbGreen
        End If

        PercentChange = Qchange / OpenPrice
        ws.Cells(TargetRow, 13).NumberFormat = "0.00%"


        ws.Cells(TargetRow, 13).Value = TotalVolume



        'Dim CorrespondingTicker As String
        'Calculate Greatest Percent Increase in Column L Percent Change
        Dim GreatestIncrease As Double
        GreatestIncrease = WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(1, 18).Value = GreatestIncrease
        ws.Cells(1, 18).NumberFormat = "0.00%"
        'CorrespondingTicker = ws.Cells(maxRow, "J").Value
        'ws.Cells(3, 17).Value = CorrespondingTicker

        'Calculate Greatest Percent Decrease in Column L Percent Change
        Dim GreatestDecrease As Double
        GreatestDecrease = WorksheetFunction.Min(ws.Range("L:L"))
        ws.Cells(2, 18).Value = GreatestDecrease
        ws.Cells(2, 18).NumberFormat = "0.00%"

        'Calculates Greatest Total Volume
        Dim GreatestTotalVolume As Double

        GreatestTotalVolume = WorksheetFunction.Max(ws.Range("M:M"))
        ws.Cells(3, 18).Value = GreatestTotalVolume
        'CorrespondingTicker = ws.Cells(maxRow, "J").Value
        'ws.Cells(3, 17).Value = CorrespondingTicker



    Next SheetName

    'MsgBox ("done")

End Sub




