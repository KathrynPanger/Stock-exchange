Attribute VB_Name = "Module2"
Sub challenge()
'CHALLENGE

'loop through the workbook
Dim Current As Worksheet
For Each Current In Worksheets

'declare variables
Dim max_percent_increase As Double
Dim max_percent_decrease As Double
Dim max_total_volume As Double
Dim max_ticker_percent As String
Dim min_ticker_percent As String
Dim max_ticker_volume As String


'create column names
Current.Cells(1, 16).Value = "Ticker"
Current.Cells(1, 17).Value = "Value"

'create row names
Current.Cells(2, 15).Value = "Greatest % Increase"
Current.Cells(3, 15).Value = "Greatest % Decrease"
Current.Cells(4, 15).Value = "Greatest Total Volume"



'PERCENT INCREASE -----------------------------------------------

'set initial max percent for testing
max_percent_increase = Current.Cells(2, 11).Value


'loop through the percent increase column
lrow = Current.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lrow
    If Current.Cells(i, 11).Value > max_percent_increase Then
        max_percent_increase = Current.Cells(i, 11).Value
        max_ticker_percent = Current.Cells(i, 9).Value
        
    Else
    End If
Next
    
  
'Put highest percent increase into a column when looping is done
Current.Cells(2, 17).Value = max_percent_increase
'put ticker associated with this value into table when looping is done
Current.Cells(2, 16).Value = max_ticker_percent

'PERCENT DECREASE -----------------------------------------------------------

'set initial min percent for testing
max_percent_decrease = Current.Cells(2, 11).Value

For i = 2 To lrow
    If Current.Cells(i, 11).Value < max_percent_decrease Then
        max_percent_decrease = Current.Cells(i, 11).Value
        min_ticker_percent = Current.Cells(i, 9).Value
        
    Else
    End If
Next
    
  
'Put highest percent decrease into a column when looping is done
Current.Cells(3, 17).Value = max_percent_decrease

'put ticker associated with this value into table when looping is done
Current.Cells(3, 16).Value = min_ticker_percent

'MAX TOTAL VOLUME-----------------------------------------------------------


'set initial max volume for testing
max_total_volume = Current.Cells(2, 12).Value

For i = 2 To lrow
    If Current.Cells(i, 12).Value > max_total_volume Then
        max_total_volume = Current.Cells(i, 12).Value
        max_ticker_volume = Current.Cells(i, 9).Value
        
    Else
    End If
Next
    
  
'Put highest ticker volume into a column when looping is done
Current.Cells(4, 17).Value = max_total_volume

'put ticker associated with this value into table when looping is done
Current.Cells(4, 16).Value = max_ticker_volume

Next

'FORMAT-----------------------------------------------------------------

'Autofit the columns
For Each Current In Worksheets
    Current.Cells.EntireColumn.AutoFit
Next

End Sub
