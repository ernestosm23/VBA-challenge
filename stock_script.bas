Attribute VB_Name = "Module1"
Sub stock_challenge()


    For Each ws In Worksheets

        'find last row
        Dim lastrow As Double
        Dim lastrow_range As String
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        lastrow_range = "J" & lastrow

        'Definining the variables:
        Dim rng As Range
        Dim condition1 As FormatCondition
        Dim condition2 As FormatCondition

        'starting point
        Set rng = ws.Range("J2", lastrow_range)

   
        'conditions
        Set greater = rng.FormatConditions.Add(xlCellValue, xlGreater, 0)
        Set lessthan = rng.FormatConditions.Add(xlCellValue, xlLess, 0)

        'Defining and setting the format to be applied for each condition
        With greater
            .Interior.ColorIndex = 4
        End With

        With lessthan
            .Interior.ColorIndex = 3
        End With

        '-------------------
        'summarizing data
        
        'adding headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        'setting variables
        Dim ticker_name As String
        Dim ticker_delta As Double
        Dim ticker_volume As Double
        Dim ticker_start As Double
        Dim ticker_end As Double
        Dim ticker_pctdelta As Double
        Dim summary_table_row As Integer

        'setting up values
        summary_table_row = 2
        ticker_volume = 0

        'summary table loop through data
        Dim i As Double
        For i = 2 To lastrow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'get summary values
                ticker_name = ws.Cells(i, 1).Value
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
                ticker_end = ws.Cells(i, 6).Value
                ticker_delta = ticker_start - ticker_end
                ticker_pctdelta = ticker_delta / ticker_start * 100
                
                'print values to summary table
                ws.Range("I" & summary_table_row).Value = ticker_name
                ws.Range("J" & summary_table_row).Value = ticker_delta
                ws.Range("K" & summary_table_row).Value = ticker_pctdelta
                ws.Range("L" & summary_table_row).Value = ticker_volume

                'next row, reset ticker volume and setting new year opening price
                summary_table_row = summary_table_row + 1
                ticker_volume = 0
                ticker_start = ws.Cells(i + 1, 3).Value

            ElseIf i = 2 Then
                'get year starting price and amount
                ticker_start = ws.Cells(i, 3).Value
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value

            Else
                'within secion, summing volume data
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
            End If

        Next i


        'applying headers for calculated values
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'finding last row for new table
        Dim newlastrow As Double
        newlastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
  
        'max, min lookups for greatest % inc, greatest % dec and highest volume - finds ticker value for each
        'adapted from wallstreetmojo.com, automateexcel.com and knowledge of index(match) functions within excel
        ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2" & ":" & "K" & newlastrow))
        ws.Cells(2, 16).Value = WorksheetFunction.Index(ws.Range("I2" & ":" & "I" & newlastrow), WorksheetFunction.Match(ws.Cells(2, 17).Value, ws.Range("K2" & ":" & "K" & newlastrow), 0))
    
        ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2" & ":" & "K" & newlastrow))
        ws.Cells(3, 16).Value = WorksheetFunction.Index(ws.Range("I2" & ":" & "I" & newlastrow), WorksheetFunction.Match(ws.Cells(3, 17).Value, ws.Range("K2" & ":" & "K" & newlastrow), 0))

        ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & newlastrow))
        ws.Cells(4, 16).Value = WorksheetFunction.Index(ws.Range("I2" & ":" & "I" & newlastrow), WorksheetFunction.Match(ws.Cells(4, 17).Value, ws.Range("L2" & ":" & "L" & newlastrow), 0))



    Next ws


End Sub

