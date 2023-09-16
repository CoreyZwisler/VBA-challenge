Sub challenge2_ws()

 'Create code for all worksheets
 For Each ws In Worksheets:

    'Declare variable for ticker as a string because the values are letters
    Dim ticker As String

    'Declare a variable for the opening value as a double because values are decimals
    Dim open_val As Double

    'Declare a variable for the closing value as a double because values are decimals
    Dim close_val As Double

    'Declare a variable for total volume as a longlong since am overflow error occured for long and assign it a 0 value
    Dim total_vol As LongLong
     total_vol = 0
    'Declare a variable to hold the yearly change as a double because the values used to calculate it are doubles
    Dim yearly_change As Double

    'Declare a variable for percentage change as a double because the values used to calculate it are doubles
    Dim percentage_change As Double

    'Create and label Summary Tables for data
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Value"

    'Create a variable and assign it a value for the Summary Table to use for loop
    Dim summary_table_row As Long
     summary_table_row = 2
    
    'Create a variable for finding the last row in the data set
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create a loop that searches through the data by ticker and posts to the summary table
    For i = 2 To lastrow
    
     'Search through the tickers by previous value to pull the opening value
     If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
    
     'Set the value for opening_val
     open_val = ws.Cells(i, 3).Value
    
     End If
     
     'Search for individual tickers
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
     'Set the value for ticker string
     ticker = ws.Cells(i, 1).Value
     
     'Print the ticker value to Summary Table
     ws.Range("i" & summary_table_row).Value = ticker
     
     'Set the value for close_val
     close_val = ws.Cells(i, 6).Value
 
     'Calculate Yearly Change
     yearly_change = close_val - open_val
 
     'Print Yearly Change to Summary Table
     ws.Range("j" & summary_table_row).Value = yearly_change
      
     'Calculate percentage change and print to summary table with a divide by 0 rule
     If open_val = 0 Then
      percentage_change = 0

     Else
     percentage_change = yearly_change / open_val
     
     End If
     
     'Print values in Summary Table and format
     ws.Range("k" & summary_table_row).Value = percentage_change
     ws.Range("k" & summary_table_row).NumberFormat = "0.00%"
     
     'Calculate the total volume
     total_vol = total_vol + ws.Cells(i, 7).Value
     
     'Print total_vol to Summary Table
     ws.Range("l" & summary_table_row).Value = total_vol
     
     'Add to summary_table_row
     summary_table_row = summary_table_row + 1
     
     'Reset total_vol
     total_vol = 0
     
     Else
     
     total_vol = total_vol + ws.Cells(i, 7).Value
     
     End If
     
    Next i
    
    'Create a new last row search for the Summary Table
    last_row_summary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Create Summary Table Loop
    For i = 2 To last_row_summary
    
     'Pull and create greatest % increase
     If ws.Cells(i, 11).Value = WorksheetFunction.Max(Range("k2:k" & summary_table_row)) Then
      ws.Cells(2, 16).Value = Cells(i, 9).Value
      ws.Cells(2, 17).Value = Cells(i, 11).Value
      ws.Cells(2, 17).NumberFormat = "0.00%"
     
     'Greatest % decrease
     ElseIf ws.Cells(i, 11).Value = WorksheetFunction.Min(Range("k2:k" & summary_table_row)) Then
      ws.Cells(3, 16).Value = Cells(i, 9).Value
      ws.Cells(3, 17).Value = Cells(i, 11).Value
      ws.Cells(3, 17).NumberFormat = "0.00%"
     
     'Greatest total volume
     ElseIf ws.Cells(i, 12).Value = WorksheetFunction.Max(Range("l2:l" & summary_table_row)) Then
      ws.Cells(4, 16).Value = Cells(i, 9).Value
      ws.Cells(4, 17).Value = Cells(i, 12).Value
      
     End If
     
     'Conditional Formatting for postitive
     If ws.Cells(i, 10).Value > 0 Then
      ws.Cells(i, 10).Interior.ColorIndex = 4
      
     'For negative
     ElseIf ws.Cells(i, 10).Value < 0 Then
      ws.Cells(i, 10).Interior.ColorIndex = 3
      
     End If
     
    Next i

 Next ws

End Sub

