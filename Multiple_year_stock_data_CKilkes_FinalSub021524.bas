Attribute VB_Name = "Module1"
'Final code for grading
Sub main_test3()
    ' Call both ticker_volume and increase_only subroutines
    ticker_volume_only_test
    increase_only_test
End Sub

Sub ticker_volume_only_test()
    Dim ws As Worksheet
    Dim last_row As Long
    Dim summary_row As Long
    Dim i As Long

    For Each ws In ThisWorkbook.Worksheets
        
        'Check if the worksheet name is "2018", "2019" and "2020"
        If ws.Name Like "20##" Then
        last_row = ws.Cells(ws.Rows.Count, "G").End(xlUp).row
        End If

        summary_row = 1 'Initialize the summary row

        ' Add headers for columns I-L
        ws.Cells(summary_row, "I").Value = "Ticker"
        ws.Cells(summary_row, "J").Value = "Yearly Change"
        ws.Cells(summary_row, "K").Value = "Percentage Change"
        ws.Cells(summary_row, "L").Value = "Volume"
        summary_row = summary_row + 1 ' Move to the next row for data

        Dim ticker As String
        Dim ticker_volume As Double
        Dim yearly_change As Double
        Dim percentage_change As Double

        'Reset variables for each new worksheet
        ticker = " "
        ticker_volume = "0"
        yearly_change = "0"
        percentage_change = "0"

        ' Loop through each row of data in the current worksheet
        For i = 2 To last_row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Output the ticker symbol to column I
                ws.Cells(summary_row, "I").Value = ticker
                ' Output yearly_change to column J
                ws.Cells(summary_row, "J").Value = yearly_change
                ' Output percentage_change to column K
                ws.Cells(summary_row, "K").Value = percentage_change
                ws.Cells(summary_row, "L").Value = ticker_volume
                summary_row = summary_row + 1
                ' Reset ticker volume for the new ticker symbol
                ticker = ws.Cells(i, 1).Value
                ticker_volume = ws.Cells(i, 3).Value
                ' Reset percentage_change and yearly_change calculations
                yearly_change = ws.Cells(i, 3) - ws.Cells(i, 6).Value
                percentage_change = (ws.Cells(i, 6) / ws.Cells(i, 3).Value) * 100
            Else
                ' Accumulate volume for the current ticker symbol
                ticker_volume = ticker_volume + ws.Cells(i, 3).Value
            End If
        Next i

        ' Output the ticker symbol and volume for the last ticker symbol in the current worksheet
        ws.Cells(summary_row, "I").Value = ticker
        ws.Cells(summary_row, "L").Value = ticker_volume
    Next ws
End Sub

Sub increase_only_test()
    Dim ws As Worksheet
    Dim last_row As Long
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As Double
    Dim max_increase_ticker As String
    Dim max_decrease_ticker As String
    Dim max_volume_ticker As String
    Dim i As Long
    Dim summary_row As Long

    For Each ws In ThisWorkbook.Worksheets
        last_row = ws.Cells(ws.Rows.Count, "G").End(xlUp).row
        max_increase = -99999999999# ' Initialize to a very small number
        max_decrease = 99999999999# ' Initialize to a very large number
        max_volume = 0
        max_increase_ticker = ""
        max_decrease_ticker = ""
        max_volume_ticker = ""
        summary_row = 2

        For i = 2 To last_row
            ' Check if the yearly change is the greatest increase
            If IsNumeric(ws.Cells(i, "J").Value) Then
                If ws.Cells(i, "J").Value > max_increase Then
                    max_increase = ws.Cells(i, "J").Value
                    max_increase_ticker = ws.Cells(i, "I").Value
                End If
            End If
            ' Check if the yearly change is the greatest decrease
           If IsNumeric(ws.Cells(i, "J").Value) Then
                If ws.Cells(i, "J").Value < max_decrease Then
                    max_decrease = ws.Cells(i, "J").Value
                    max_decrease_ticker = ws.Cells(i, "I").Value
                End If
            End If
            ' Check if the total volume is the greatest
            If IsNumeric(ws.Cells(i, "L").Value) Then
                If ws.Cells(i, "L").Value > max_volume Then
                    max_volume = ws.Cells(i, "L").Value
                    max_volume_ticker = ws.Cells(i, "I").Value
                End If
            End If
        Next i

        ' Output the results on the current worksheet
        ws.Cells(summary_row, "N").Value = "Greatest % Increase"
        ws.Cells(summary_row + 1, "N").Value = "Greatest % Decrease"
        ws.Cells(summary_row + 2, "N").Value = "Greatest Total Volume"
        ws.Cells(1, "O").Value = "Ticker"
        ws.Cells(1, "P").Value = "Value"
        ws.Cells(summary_row, "O").Value = max_increase_ticker
        ws.Cells(summary_row + 1, "O").Value = max_decrease_ticker
        ws.Cells(summary_row + 2, "O").Value = max_volume_ticker
        ws.Cells(summary_row, "P").Value = max_increase & "%"
        ws.Cells(summary_row + 1, "P").Value = max_decrease & "%"
        ws.Cells(summary_row + 2, "P").Value = max_volume
    Next ws
End Sub
