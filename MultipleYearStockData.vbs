'Create a script that loops through all the stocks for each quarter and outputs the following information:

'The ticker symbol

'Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

'The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

Sub MultipleYearStockData()
    
    Dim Q1 As Worksheet
    Dim Q2 As Worksheet
    Dim Q3 As Worksheet
    Dim Q4 As Worksheet
    Dim i As Integer
  

    ' Set the worksheet variables to the corresponding sheets
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
    
    'Go through all tickers in Q1 (column A) and list all unique tickers to colum Ticker (J)
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
    
    For InputRow = 2 To LastRow
        
        
        If ws.Cells(InputRow, 1).Value <> Ticker Then
            ' Ticker changes from InputRow - 1 to InputRow
            
            ' First, "close" the previous Ticker information
            
            If InputRow > 2 Then
                ' Make sure there is previous Ticker
                
                ClosePrice = ws.Cells(InputRow - 1, 6).Value
                Qchange = ClosePrice - OpenPrice
                ws.Cells(TargetRow, 11).Value = Qchange
            End If
            
            ' Then, "open" the next Ticker information

            Ticker = ws.Cells(InputRow, 1)
            
            TargetRow = TargetRow + 1
            ws.Cells(TargetRow, 10).Value = Ticker
            OpenPrice = ws.Cells(InputRow, 3).Value
            
        End If
        
    Next InputRow
    
    ' Finally, "close" the last Ticker
    ClosePrice = ws.Cells(InputRow - 1, 6).Value
    Qchange = ClosePrice - OpenPrice
    ws.Cells(TargetRow, 11).Value = Qchange
    
    'MsgBox (LastRow)
    
     
End Sub



