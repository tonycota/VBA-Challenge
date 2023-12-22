Sub module_two_challenge()

'define all of the functions you will be using for the script

Dim ws As Worksheet
Dim opened As Double
Dim vol As Double
Dim total As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim s_value As Double
Dim lastrow As Long
Dim sum As Double
Dim i As Long


'set up for loop to run through ALL scripts
'looping through all sheets source in README
For Each ws In Worksheets

    'add value to new header cells, adding value via brackets source in README
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'mark the starting values of listed variables to 0 for the loop
    
    opened = 2
    s_value = 0
    total = 0
    sum = 2
    
    'find the last row using xl up function, source in README
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set i to loop through rows 2 to last row
    For i = 2 To lastrow
        'conditional; check to see if previous cells are different from current cells
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            total = total + ws.Cells(i, 7).Value 'hold results
            
                If total = 0 Then
                    'populate results
                    ws.Range("i" & 2 + s_value).Value = ws.Cells(i, 1).Value
                    ws.Range("j" & 2 + s_value).Value = 0
                    ws.Range("k" & 2 + s_value).Value = "%" & 0
                    ws.Range("l" & 2 + s_value).Value = 0
                
               Else
                    
                    If ws.Cells(opened, 3) = 0 Then
                        For vol = opened To i
                            If ws.Cells(vol, 3).Value <> 0 Then
                                opened = vol
                            End If
                        Next vol
                    End If
                    
                    'assign to yearly and percent change variables
                    yearly_change = (ws.Cells(i, 6) - ws.Cells(opened, 3))
                    percent_change = Round((yearly_change / ws.Cells(opened, 3) * 100), 2)
                    
                    opened = i + 1
                    
                    'populate the spreadsheet
                    ws.Range("i" & 2 + s_value).Value = ws.Cells(i, 1).Value
                    'round function source in README file
                    ws.Range("j" & 2 + s_value).Value = Round(yearly_change, 2)
                    ws.Range("k" & 2 + s_value).Value = "%" & percent_change
                    ws.Range("l" & 2 + s_value).Value = total
                    
                End If

                    
                    'conditonal format the [J:J] column accordingly for yearly change
                    
                    If ws.Range("j" & 2 + s_value).Value > 0 Then
                        ws.Range("j" & 2 + s_value).Interior.ColorIndex = 4
                    ElseIf ws.Range("j" & 2 + s_value).Value < 0 Then
                        ws.Range("j" & 2 + s_value).Interior.ColorIndex = 3
                    Else
                        ws.Range("j" & 2 + lastrow).Interior.ColorIndex = 0
                    End If
                    
                    'condtionally format the [K:K] column accordingly for percent change
                    
                    If ws.Range("k" & 2 + s_value).Value > 0 Then
                        ws.Range("k" & 2 + s_value).Interior.ColorIndex = 4
                    ElseIf ws.Range("k" & 2 + s_value).Value < 0 Then
                        ws.Range("k" & 2 + s_value).Interior.ColorIndex = 3
                    Else
                        ws.Range("k" & 2 + lastrow).Interior.ColorIndex = 0
                    End If
            
                'reset values for next loop
                total = 0
                yearly_change = 0
                s_value = s_value + 1
            
            Else
                total = total + ws.Cells(i, 7).Value
                
        End If
                                 

    Next i
     
    
    'find the min and max of columns K and L, source in README
    ws.[Q2].Value = "%" & WorksheetFunction.Max(Range("k2:k" & lastrow)) * 100
    ws.[Q3].Value = "%" & WorksheetFunction.Min(Range("k2:k" & lastrow)) * 100
    ws.[Q4].Value = WorksheetFunction.Max(Range("l2:l" & lastrow)) 'greatest total volume
    
    'attach the values to the index functions
    Index1 = WorksheetFunction.Match(WorksheetFunction.Max(Range("k2:k" & lastrow)), Range("k2:k" & lastrow), 0)
    Index2 = WorksheetFunction.Match(WorksheetFunction.Min(Range("k2:k" & lastrow)), Range("k2:k" & lastrow), 0)
    Index3 = WorksheetFunction.Match(WorksheetFunction.Max(Range("l2:l" & lastrow)), Range("l2:l" & lastrow), 0)
    
    ws.[P2].Value = ws.Cells(Index1 + 1, 9).Value
    ws.[P3].Value = ws.Cells(Index2 + 1, 9).Value
    ws.[P4].Value = ws.Cells(Index3 + 1, 9).Value
    
 'close the entire loop and move on to the next sheet
Next ws


End Sub
