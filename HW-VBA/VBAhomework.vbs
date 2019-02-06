Sub total_stock_volume()
Dim ws As Worksheet
Dim percentchange As Double
Dim max As Variant
Dim min As Variant
Dim open_value As Variant
Dim close_value As Variant
Dim myarray1() As Variant                                               ' Array to store the results from percent change
Dim max_volume As Double                                                'Set a variable to compare the max stock volume
Dim ticker_name As Variant                                              'Set a varibale for holding ticker_name/name
Dim counter As Variant
Dim yearly_change As String                                             'Set a variable for holding yearly_change value
Dim percent_change As String                                            'Set a variable for holding percent_change value
Dim opening_cost As Variant                                             'Set a variable for holding the opening cost of a ticker in a year
Dim closing_cost As Variant                                            'Set a variable for holding the closing cost of a ticker in a year
Dim ts_tracker As Integer                                               'Keep track of the ticker_name location/ summary table row tracker
Dim lastrow As Variant                                                  'Set a variable to find last row
Dim endrow As Variant                                                   'Set a variable to store end row for array
For Each ws In Worksheets
    ws.Range("I1").Value = "Ticker"                                     'Adding a new column name for column I
    ws.Range("J1").Value = "Yearly Change"                              'Adding a new column name for column J
    ws.Range("K1").Value = "Percent Change"                             'Adding a new column name for column K
    ws.Range("L1").Value = "Total stock volume"                         'Adding a new column name for column L
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ts_tracker = 2
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row                   'Find the last non-blank row
    For i = 2 To lastrow                                                'Loop through all the ticker_name names
        On Error Resume Next
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then        'Check if we're still within the same ticker_name
            'Call percent increase and decrease here
            open_value = ws.Cells(i - counter, 6).Value
            close_value = ws.Cells(i, 6).Value
            percentchange = (close_value - open_value) / open_value
            ticker_name = ws.Cells(i, 1).Value                          'Set the ticker_name
            ReDim Preserve myarray1(ts_tracker - 1)
            myarray1(ts_tracker - 1) = percentchange
            ts_volume = ts_volume + ws.Cells(i, 7).Value                'Add to the total stock volume
            ws.Range("I" & ts_tracker).Value = ticker_name              'Print the ticker_name to the summary table
            ws.Range("K" & ts_tracker).Value = percent_change           'Print the percent change to the summary table
            ws.Range("L" & ts_tracker).Value = ts_volume                'Print the total stock volume to the summary table
            closing_cost = ws.Cells(i, 6).Value                         'Get the closing cost
            If ws.Cells(ts_tracker, 9).Value = ws.Cells(i, 1).Value Then 'Condition to get the opening cost
                opening_cost = ws.Cells(i, 3).Value                     'Assigning opening cost to variable
            End If
            ws.Range("J" & ts_tracker).Value = close_value - open_value            'Print the yearly change to the summary table
            If ws.Range("J" & ts_tracker).Value > 0 Then
                ws.Range("J" & ts_tracker).Interior.ColorIndex = 4
            Else
                ws.Range("J" & ts_tracker).Interior.ColorIndex = 3
            End If
            ws.Range("J" & ts_tracker).NumberFormat = "0.0000000000"
            ws.Range("K" & ts_tracker).Value = (close_value - open_value) / open_value    'Print the percent chanage to the summary table
            ws.Range("K" & ts_tracker).Style = "Percent"
            ws.Range("K" & ts_tracker).NumberFormat = "0.00%"
            ts_tracker = ts_tracker + 1                                 'Add one to the summary table row tracker
            ts_volume = 0                                               'Reset the ts_volume for ticker_name
            counter = 0
        Else                                                            'If the cell immediately following a row is the same ticker_name
            ts_volume = ts_volume + ws.Cells(i, 7).Value                'Add to the ts_volume
            counter = counter + 1
        End If
     Next i
    Columns("A:Q").EntireColumn.AutoFit
      
    For j = 0 To ts_tracker
        ws.Cells(j, 13).Value = myarray1(j)
        ws.Range("M" & j).Style = "Percent"
        ws.Range("M" & j).NumberFormat = "0.00%"
    Next j

endrow = ws.Cells(Rows.Count, "M").End(xlUp).Row
min = 1000
max = -1000
    For k = 0 To endrow
        If min > ws.Cells(k, 13).Value Then
            min = ws.Cells(k, 13).Value
            ws.Range("Q3").Value = ws.Cells(k, 9).Value
            ws.Range("R3").Value = min
            ws.Range("R3").Style = "Percent"
            ws.Range("R3").NumberFormat = "0.00%"
        End If
        If max < ws.Cells(k, 13).Value Then
            max = ws.Cells(k, 13).Value
            ws.Range("Q2").Value = ws.Cells(k, 9).Value
            ws.Range("R2").Value = max
            ws.Range("R2").Style = "Percent"
            ws.Range("R2").NumberFormat = "0.00%"
        End If
    Next k

    max_volume = 0
    endrow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    For i = 2 To endrow
        If max_volume < ws.Cells(i, 12).Value Then
           max_volume = ws.Cells(i, 12).Value
           ws.Range("Q4").Value = ws.Cells(i, 9).Value
           ws.Range("R4").Value = max_volume
           ws.Range("R4").NumberFormat = "0"
        End If
    Next i
ws.Columns("M").Delete
ws.Columns("O:Q").EntireColumn.AutoFit
Next ws
Columns("O:Q").EntireColumn.AutoFit
End Sub