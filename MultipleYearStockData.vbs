'Create a script that loops through all the stocks For each quarter And outputs the following information:

'The ticker symbol

'Quarterly change from the opening price at the beginning of a given quarter To the closing price at the end of that quarter.

'The percentage change from the opening price at the beginning of a given quarter To the closing price at the end of that quarter.

Sub MultipleYearStockData()

    Dim Q1 As Worksheet
    Dim Q2 As Worksheet
    Dim Q3 As Worksheet
    Dim Q4 As Worksheet
    Dim i As Integer


    ' Set the worksheet variables To the corresponding sheets
    Set Q1 = ThisWorkbook.Sheets("Q1")
    Set Q2 = ThisWorkbook.Sheets("Q2")
    Set Q3 = ThisWorkbook.Sheets("Q3")
    Set Q4 = ThisWorkbook.Sheets("Q4")


  
    Q1.Cells(1, 10).Value = "Ticker"
    Q1.Cells(1, 11).Value = "Quarterly Change"
    Q1.Cells(1, 12).Value = "Percent Change"
    Q1.Cells(1, 13).Value = "Total Stock Volume"


    Dim ws As Worksheet
    Dim LastRow As Long

    Set ws = ThisWorkbook.Sheets("Q1")
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row


    Dim InputRow As Long
    Dim OutputRow As Long
    Dim Ticker As String
    Dim TargetRow As Long

    TargetRow = 1

    'Go through all tickers in Q1 (column A) And list all unique tickers To colum Ticker (J)
    'For InputRow = 2 To LastRow
    'If ws.Cells(InputRow, 1).Value <> Ticker Then
    'Ticker = ws.Cells(InputRow, 1)
    'ws.Cells(TargetRow, 10).Value = Ticker
    'TargetRow = TargetRow + 1
    'End If
    'Next InputRow


    Dim Qchange As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PercentChange As Double
    Dim RoundedPercentage As Double
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

                ' Calculate percent change
                PercentChange = Qchange / OpenPrice * 100
                RoundedPercentage = Round(PercentChange, 2)
                ws.Cells(TargetRow, 12).Value = RoundedPercentage

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

    PercentChange = Qchange / OpenPrice * 100
    RoundedPercentage = Round(PercentChange, 2)
    ws.Cells(TargetRow, 12).Value = RoundedPercentage


    ws.Cells(TargetRow, 13).Value = TotalVolume



End Sub



