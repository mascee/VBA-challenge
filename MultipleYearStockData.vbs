'Create a script that loops through all the stocks For each quarter And outputs the following information:

'The ticker symbol

'Quarterly change from the opening price at the beginning of a given quarter To the closing price at the end of that quarter.

'The percentage change from the opening price at the beginning of a given quarter To the closing price at the end of that quarter.
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".


Sub MultipleYearStockData()

    Dim SheetName As Variant
    Dim SheetNames As Variant
    
    'Declare names for different sheets
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
        Dim TotalVolume As Double
        Dim Volume As Double

        For InputRow = 2 To LastRow


            If ws.Cells(InputRow, 1).Value <> Ticker Then
                
                'Skips first row
                If InputRow > 2 Then
                    ClosePrice = ws.Cells(InputRow - 1, 6).Value
                    Qchange = ClosePrice - OpenPrice
                    ws.Cells(TargetRow, 11).Value = Qchange
                    'Format color
                    If Qchange < 0 Then
                        ws.Cells(TargetRow, 11).Interior.Color = vbRed
                    ElseIf Qchange > 0 Then
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
            
            'Calculate total volume
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
        ElseIf Qchange > 0 Then
            ws.Cells(TargetRow, 11).Interior.Color = vbGreen
        End If

        PercentChange = Qchange / OpenPrice
        ws.Cells(TargetRow, 13).NumberFormat = "0.00%"


        ws.Cells(TargetRow, 13).Value = TotalVolume



        'Loop to find Max Percent Increase
        Dim LastRow2 As Long
        Dim PercentRow As Long
        Dim GetTicker As String
        LastRow2 = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
        'Greatest Increase
        Dim GIncrease As Double
        GIncrease = ws.Cells(2, 12).Value
         For PercentRow = 2 To LastRow2
            If ws.Cells(PercentRow, 12).Value > GIncrease Then
                GIncrease = ws.Cells(PercentRow, 12).Value
                ws.Cells(1, 18).Value = GIncrease
        'Get corresponding ticker
        GetTicker = ws.Cells(PercentRow, 10).Value
        ws.Cells(1, 17).Value = GetTicker
            End If
        Next PercentRow

                
        
        'Greatest Decrease
        Dim GDecrease As Double
        GDecrease = ws.Cells(2, 12).Value
         For PercentRow = 2 To LastRow2
            If ws.Cells(PercentRow, 12).Value < GDecrease Then
                GDecrease = ws.Cells(PercentRow, 12).Value
                ws.Cells(2, 18).Value = GDecrease
                'Get corresponding ticker
                GetTicker = ws.Cells(PercentRow, 10).Value
                ws.Cells(2, 17).Value = GetTicker
            End If
         Next PercentRow
        
        'Loops to find Greatest Volume
        Dim GTotalVolume As Double
        Dim TotalVolumeRow As Double
        Dim LastRow3 As Long
        LastRow3 = ws.Cells(ws.Rows.Count, 13).End(xlUp).Row
        GTotalVolume = ws.Cells(2, 13).Value
            For TotalVolumeRow = 2 To LastRow3
                If ws.Cells(TotalVolumeRow, 13).Value > GTotalVolume Then
                    GTotalVolume = ws.Cells(TotalVolumeRow, 13).Value
                    ws.Cells(3, 18).Value = GTotalVolume
                    'Get corresponding ticker
                    GetTicker = ws.Cells(TotalVolumeRow, 10).Value
                    ws.Cells(3, 17) = GetTicker
                End If
            Next TotalVolumeRow
        

    Next SheetName


End Sub





