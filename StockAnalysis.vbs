Attribute VB_Name = "Module1"
' Ramamurthy Sundar
  
' This subroutine performs a stock analysis of thousands of stocks and
' posts the results on the given spreadsheets.
' This sub works on every worksheet in the workbook.
' The summary data includes, total stock volume, the percentage change over the course
' of the year, and the total change of the stock over the course of the year.  There is
' also a summary table (table 3), which tells us which stocks saw the greatest changes
' over the course of the year.

Sub StockAnalysis()
    ' Define arrays
    Dim ticker_array(10000) As String
    Dim volume_array(10000) As Variant
    Dim yearly_change_array(10000) As Double
    Dim percentage_change_array(10000) As Double
    
    ' Define key variables
    Dim num_rows As Long
    Dim array_size As Integer
    Dim running_total As Variant
    Dim beg_open_price As Double
    Dim end_close_price As Double
    Dim temp_percentage_change As Double
    Dim temp_yearly_change As Double
    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_volume As Variant
    Dim ws As Worksheet
    
    ' loop through each worksheet in the workbook
    For Each ws In ActiveWorkbook.Worksheets
        'determine number of rows
        num_rows = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' initialize with data from the first row of the table
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ticker_array(0) = ws.Cells(2, 1).Value
        running_total = ws.Cells(2, 7).Value
        beg_open_price = ws.Cells(2, 3).Value
        array_size = 1
        
        ' initialize headers in table 3
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest total volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'loop through all the rows and populate the
        For i = 3 To num_rows
            ' Case 1: The ticker symbol is not unique
            If (ticker_array(array_size - 1) = ws.Cells(i, 1).Value) Then
                running_total = running_total + ws.Cells(i, 7).Value
                ' Edge case - End of worksheet
                If (i = num_rows) Then
                    ' assign close price
                    end_close_price = ws.Cells(i, 6).Value
                    ' calculate the percentage change and yearly change
                    ' selection is used to prevent a divide by zero error
                    If (beg_open_price = 0) Then
                        temp_percentage_change = 0
                        temp_yearly_change = end_close_price - beg_open_price
                    Else
                        temp_percentage_change = end_close_price / beg_open_price * 100 - 100
                        temp_yearly_change = end_close_price - beg_open_price
                    End If
                    ' update the volume_array
                    volume_array(array_size - 1) = running_total
                    yearly_change_array(array_size - 1) = temp_yearly_change
                    percentage_change_array(array_size - 1) = temp_percentage_change
                End If
            ' Case 2: The ticker symbol is unique
            Else
                ' assign close price
                end_close_price = ws.Cells(i - 1, 6).Value
                ' calculate the percentage change and yearly change
                ' selection is used to prevent a divide by zero error
                If (beg_open_price = 0) Then
                    temp_percentage_change = 0
                    temp_yearly_change = end_close_price - beg_open_price
                Else
                    temp_percentage_change = end_close_price / beg_open_price * 100 - 100
                    temp_yearly_change = end_close_price - beg_open_price
                End If
                ' update the arrays
                ticker_array(array_size) = ws.Cells(i, 1).Value
                volume_array(array_size - 1) = running_total
                yearly_change_array(array_size - 1) = temp_yearly_change
                percentage_change_array(array_size - 1) = temp_percentage_change
                ' update variables
                beg_open_price = ws.Cells(i, 3)
                array_size = array_size + 1
                running_total = ws.Cells(i, 7)
            End If
        Next i
        
        ' variables needed for table 3
        max_percent = Application.WorksheetFunction.Max(percentage_change_array)
        min_percent = Application.WorksheetFunction.Min(percentage_change_array)
        max_volume = Application.WorksheetFunction.Max(volume_array)
        ws.Cells(2, 17).Value = max_percent
        ws.Cells(3, 17).Value = min_percent
        ws.Cells(4, 17).Value = max_volume
        
        ' populate the spreadsheet from the values stored in the arrays
        For i = 0 To array_size - 1
            ' populate values of arrays in table 2
            ws.Cells(i + 2, 9).Value = ticker_array(i)
            ws.Cells(i + 2, 10).Value = yearly_change_array(i)
            ws.Cells(i + 2, 11).Value = Str(percentage_change_array(i)) + "%"
            ws.Cells(i + 2, 12).Value = volume_array(i)
            
            ' conditional formatting for yearly change columns
            If (yearly_change_array(i) <= 0) Then
                ws.Cells(i + 2, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i + 2, 10).Interior.ColorIndex = 4
            End If
            
            ' populates values in table 3 (summary data of the stocks)
            If (percentage_change_array(i) = max_percent) Then
                ws.Cells(2, 16) = ticker_array(i)
            End If
            If (percentage_change_array(i) = min_percent) Then
                ws.Cells(3, 16) = ticker_array(i)
            End If
            If (volume_array(i) = max_volume) Then
                ws.Cells(4, 16) = ticker_array(i)
            End If
        Next i
    Next
End Sub


