Attribute VB_Name = "Module1"
Sub StockAnalysis()

    ' Define Variables
    Dim ticker As String
    Dim Total_Volume As Double
    Dim Summary_Table As Integer
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Yearly_Percent_Change As Double
    Dim Max_Total_Volume As Double
    Dim Max_Yearly_Percent_Change As Double
    Dim Max_Decrease_Yearly_Percent_Change As Double
    Dim Max_Decrease_Ticker As String

    ' Initialize maximum values to zero
    Max_Total_Volume = 0
    Max_Yearly_Percent_Change = 0
    Max_Decrease_Yearly_Percent_Change = 0
    Max_Decrease_Ticker = ""

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        Total_Volume = 0
        Summary_Table = 2 ' Reset Summary_Table to 2 for each new sheet

        Last_Row = ws.Cells(Rows.Count, "A").End(xlUp).Row

        ' Create Summary Table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Yearly Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"

        ' Looping through rows
        For i = 2 To Last_Row
                
            ' Find cells of unique tickers
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                
                ' Initialize Open_Price for the new ticker
                Open_Price = ws.Cells(i, 3).Value

                ' Find Close_Price for the last row of the current ticker
                Close_Price = ws.Cells(i, 6).Value

                ' Calculating Yearly Change
                Yearly_Change = Close_Price - Open_Price

                ' Calculate Percent Change
                If Yearly_Change <> 0 And Open_Price <> 0 Then
                    Yearly_Percent_Change = Yearly_Change / Open_Price
                Else
                    Yearly_Percent_Change = 0
                End If

                ' Calculating Total Volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value

                ' Input Values on Summary Table
                ws.Range("I" & Summary_Table).Value = ticker
                ws.Range("J" & Summary_Table).Value = Yearly_Change
                ws.Range("K" & Summary_Table).Value = Yearly_Percent_Change
                ws.Range("L" & Summary_Table).Value = Total_Volume
                Summary_Table = Summary_Table + 1

                ' Update maximum values
                If Total_Volume > Max_Total_Volume Then
                    Max_Total_Volume = Total_Volume
                End If

                If Yearly_Percent_Change > Max_Yearly_Percent_Change Then
                    Max_Yearly_Percent_Change = Yearly_Percent_Change
                End If

                If Yearly_Percent_Change < Max_Decrease_Yearly_Percent_Change Then
                    Max_Decrease_Yearly_Percent_Change = Yearly_Percent_Change
                    Max_Decrease_Ticker = ticker
                End If

                ' Reset Total Volume for the next ticker
                Total_Volume = 0
            Else
                ' Continue accumulating Total Volume for the same ticker
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            End If
        Next i

        ' Display red for negative, green for positive, and no color for no change Yearly_Change in values
        For i = 2 To Last_Row
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 0
            End If

            ws.Cells(i, 10).Style = "Currency"
            ws.Cells(i, 11).NumberFormat = "0.00%"
        Next i

        ' Display the maximum values for each sheet
        ws.Cells(2, 14).Value = "Max Total Volume"
        ws.Cells(2, 15).Value = Max_Total_Volume
        ws.Cells(3, 14).Value = "Max Yearly Percent Change"
        ws.Cells(3, 15).Value = Max_Yearly_Percent_Change
        ws.Cells(4, 14).Value = "Max Decrease Yearly Percent Change"
        ws.Cells(4, 15).Value = Max_Decrease_Yearly_Percent_Change
        ws.Cells(5, 14).Value = "Max Decrease Ticker"
        ws.Cells(5, 15).Value = Max_Decrease_Ticker

        ' Reset maximum values for the next sheet
        Max_Total_Volume = 0
        Max_Yearly_Percent_Change = 0
        Max_Decrease_Yearly_Percent_Change = 0
        Max_Decrease_Ticker = ""
    Next ws

End Sub

