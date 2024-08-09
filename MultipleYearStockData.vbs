'Create a script that loops through all the stocks for each quarter and outputs the following information:

'The ticker symbol

'Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

'The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

Sub MultipleYearStockData()
    'Dim MultipleYearStockData As String
    'Dim Q1 As Worksheet
    'Dim Q2 As Worksheet
    'Dim Q3 As Worksheet
    'Dim Q4 As Worksheet
    
    Dim i As Integer
    'ThisWorkbook.Sheets("Sheet1").Range("J1").Value = "Quarterly Change"
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Quarterly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"
    
    'ws.Cells(1, 10).Value = "Quarterly Change"
    'Dim Ticker As String
    'Ticker = ws.Cells(1, 1).Value
    'ws.Cells(1, 10).Value = ws.Cells(1, 1).Value
    'ws.Range(1, 10) = ws.Range("A1")
    
    'for i = 2 to
    
    Dim ws As Worksheet
    Dim LastRow As Long
    
    Set ws = ThisWorkbook.Worksheets("Q1")
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim InputRow As Long
    Dim OutputRow As Long
    Dim Ticker As String
    
    TargetRow = 2
    
    For InputRow = 2 To LastRow
        If ws.Cells(InputRow, 1).Value <> Ticker Then
            Ticker = ws.Cells(InputRow, 1)
            ws.Cells(TargetRow, 10).Value = Ticker
            TargetRow = TargetRow + 1
        End If
    Next InputRow
    
    

    
    MsgBox (LastRow)
    
    
    
End Sub



