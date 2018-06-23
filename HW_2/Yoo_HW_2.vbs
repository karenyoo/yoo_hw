Sub easy()

For Each ws In Worksheets

Dim ticker_name As String
Dim ticker_total As Double

'Set an initial variable for holding the total per ticker
Dim summary_table_row As Integer
summary_table_row = 2

'Set a range for the last row
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set header for the summary table
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"

    For i = 2 To last_row

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

            'Set the ticker name
            ticker_name = ws.Cells(i, 1).Value
            ticker_total = ticker_total + ws.Cells(i, 7).Value

            'Print the ticker name and total in the summary table
            ws.Range("I" & summary_table_row).Value = ticker_name
            ws.Range("J" & summary_table_row).Value = ticker_total

            summary_table_row = summary_table_row + 1

            'Reset the ticker total
            ticker_total = 0

    Else
        ticker_total = ticker_total + ws.Cells(i, 7).Value
    
        End If

    Next i

Next ws

End Sub
