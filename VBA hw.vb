Option Explicit

Sub hw()

'Dim variables
'Dim worksheet
Dim ws As Worksheet
'Total counter for each ticker
'Row counter for data table
'Row counter for output table
Dim total, row_data, row_output As Integer
'Yearly Change for each ticker
'Percent Change for each ticker
'Closing price end of year
'Opening price end of year
Dim yearly_change, percent_change, closing_price, opening_price As Double
'*BONUS*
'Dim each variables to display in bonus summary table
Dim max_increase_ticker, max_decrease_ticker, max_total_ticker As String
Dim max_increase_value, max_decrease_value, max_total_value As Double

    'Loop through each worksheets
    For Each ws In Worksheets
        total = 0
        row_output = 2

        'Summary Table outline
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        '*BONUS*
        'Bonus table header
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'First ticker opening price (no ticker to change from last one)
        opening_price = ws.Range("C2").Value

        'Loop through all the ticker rows
        For row_data = 2 To ws.Range("A1").End(xlDown).Row

            'Is next row the same ticker?

            'If different ticker...
            If ws.Cells(row_data, 1).Value <> ws.Cells(row_data + 1, 1).Value Then
                'Copy ticker to column I
                ws.Cells(row_output, 9).Value = ws.Cells(row_data, 1).Value
                'Save total to column L
                ws.Cells(row_output, 12).Value = total
                'Then reset total counter to 0
                total = 0
                'Define closing_price
                closing_price = ws.Cells(row_data, 6).Value

                    'Division by zero error: need If statement to list change as 0
                    If opening_price = 0 Then
                        percent_change = 0
                    Else
                        'Formula for Percent Change
                        percent_change = ((closing_price - opening_price) / opening_price)
                    End If

                'Formula for Yearly Change
                yearly_change = closing_price - opening_price
                'Save yearly_change to column J
                ws.Cells(row_output, 10).Value = yearly_change
                'Save percent_change to column K
                ws.Cells(row_output, 11).Value = percent_change
                'Add opening_price for next ticker
                opening_price = ws.Cells(row_data + 1, 3).Value
                'Add 1 row to row_output counter
                row_output = row_output + 1

            'If same ticker...
            Else
                'Add vol from column G to total
                total = total + ws.Cells(row_data, 7).Value
            End If

            'Conditional formatting: positive in green & negative in red
            If ws.Cells(row_data, 10).Value > 0 Then
                'Color cell green if positive
                ws.Cells(row_data, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(row_data, 10).Value < 0 Then
                'Color cell red if negative
                ws.Cells(row_data, 10).Interior.ColorIndex = 3
            Else
                'So we don't keep coloring the blank cells
                ws.Cells(row_data, 10).Interior.ColorIndex = 0  
            End If

            'Format percent_change as percentage
            ws.Cells(row_data, 11).NumberFormat = "0.00%"

            '*BONUS*
            'Define max values to fill in the bonus table
            If (percent_change > max_increase_value) Then
                max_increase_value = percent_change
                'Pull the corresponding ticker name from column A
                max_increase_ticker = ws.Cells(row_data, 1).Value
            'Define min values for bonus table
            ElseIf(percent_change < max_decrease_value) Then
                max_decrease_value = percent_change
                'Pull the corresponding ticker name from column A
                max_decrease_ticker = ws.Cells(row_data, 1).Value
            End If

            If (total > max_total_value) Then
                max_total_value = total
                'Pull the corresponding ticker name from column A
                max_total_ticker = ws.Cells(row_data, 1).Value
            End If
        Next row_data

            'Bonus table content
            ws.Range("P2").Value = max_increase_ticker
            ws.Range("P3").Value = max_decrease_ticker
            ws.Range("P4").Value = max_total_ticker
            ws.Range("Q2").Value = max_increase_value
            ws.Range("Q3").Value = max_decrease_value
            ws.Range("Q4").Value = max_total_value

            'Format max and min value in table as percentage
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
    Next ws
End Sub
