Sub alphabetical_testing()

'' ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    'Inserting Ticker & Total Stock Vol.:
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"
'Ticker & TSV Variables:
    Dim Ticker As String
    Dim Total_Stock_Volume As LongLong
    Total_Stock_Volume = 0
'Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'Keep Track of Ticker in the table:
    Dim Table_Row As Integer
    Table_Row = 2
'Going through Stocks:
    For i = 2 To LastRow
'Where are you now_condition:
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker = ws.Cells(i, 1).Value
'Total Stock Volume Operation:
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
'Show Total_Row:
    ws.Range("I" & Table_Row).Value = Ticker
'Show Total
    ws.Range("J" & Table_Row).Value = Total_Stock_Volume
'Add 1 to Table_Row:
    Table_Row = Table_Row + 1
'Ticker must be reset
    Ticker = 0
Else
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value



End If
Next i
Next ws

End Sub

