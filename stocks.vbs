Attribute VB_Name = "Module11"

Sub stocks()
'WRAP EVERYTHING IN A LOOP FOR ALL WORKSHEETS

' Declare Current as a worksheet object variable.
Dim Current As Worksheet

' Loop through all of the worksheets in the active workbook.
For Each Current In Worksheets



'DEFINE VARIABLES

    Dim lastRow As Integer
    Dim lrow As Long
    Dim lcolumn As Long
    lrow = Current.Cells(Rows.Count, 1).End(xlUp).Row
    lCol = Current.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
    Dim tickersum As Double
    tickersum = 0
    Dim summary_table_row As Double
    Dim tickername As String
    Dim open_value As Double
    Dim close_value As Double
    Dim percent_change As Double
    
    
    'set initial values
    summary_table_row = 2
    
    'get the open date of the first ticker
    open_value = Current.Cells(2, 3).Value
    
'CREATE TABLE HEADERS
Current.Cells(1, 9).Value = "Ticker"
Current.Cells(1, 10).Value = "Yearly Change"
Current.Cells(1, 11).Value = "Percent Change"
Current.Cells(1, 12).Value = "Total Stock Volume"

    
'LOOP THROUGH TICKER NAMES
    For i = 2 To lrow
        If Current.Cells(i + 1, 1) <> Cells(i, 1).Value Then
            
            'GET THE TICKER VOLUME
            
            'set the ticker name
            tickername = Cells(i, 1).Value
            
            'add to the ticker volume total
            tickersum = tickersum + Cells(i, 7).Value
            
            'print the name of the current ticker to the summary table
            Current.Cells(summary_table_row, 9).Value = Cells(i, 1)
            
            'print the ticker volume to the summary table
            Current.Cells(summary_table_row, 12).Value = tickersum
            
            'reset the ticker sum
            tickersum = 0
            
            'GET THE YEARLY CHANGE
            
            'get the close value of the current ticker
            close_value = Current.Cells(i, 6).Value
           
            
            'get the change between open and close value of current ticker
            changed_value = open_value - close_value
            
            'put the changed value in the summary table
            Current.Cells(summary_table_row, 10).Value = changed_value
            
            'get percent change betweeen open and close value of current ticker
                If close_value = 0 Then
                    percent_change = 0
                Else
                    percent_change = ((open_value / close_value) * 100) - 100
                    
                End If
            
            'put the percent change in the summary table
            Current.Cells(summary_table_row, 11).Value = percent_change
            
            
            'get the open value of the next ticker
            open_value = Current.Cells(i + 1, 3).Value
            
            
            
            'FINISH LOOP! Add one to the summary table
            summary_table_row = summary_table_row + 1
        
        Else
            'add to the tickersum
            tickersum = tickersum + Current.Cells(i, 7).Value
        End If
        
    Next

'COLOR THE CELLS

    'declare cell number variable

    Dim j As Long
    
    'loop through cells

        For j = 2 To lrow
            If Current.Cells(j, 11) <= 0 Then
                Current.Cells(j, 11).Interior.ColorIndex = 3
            Else
                Current.Cells(j, 11).Interior.ColorIndex = 4
            End If
        Next
        
           
            
    
'CLEAN THE WORKSHEET

'Autofit the columns on every worsheet

Current.Cells.EntireColumn.AutoFit
    
Next

End Sub
