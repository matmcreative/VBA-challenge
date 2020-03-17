Sub Stock_Analysis()
     
    'Set initial variables
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Yearly_Change As Single
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Start_Value As Long
    Dim LastRow As Long
    Dim i As Long
    Dim j As Long
    
    'Loop through all of the worksheets in the active workbook
    For Each ws In Worksheets
    
    'Set values for each worksheet
        j = 0
        Total_Stock_Volume = 0
        Yearly_Change = 0
        Start_Value = 2
 
        'Add headers for results
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'get the row number of the last row with data
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            For i = 2 To LastRow

                'Check if we are still within the same stock ticker if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    'Add to the Volume Total
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                    
                    'Handle zero total volume
                    If Total_Stock_Volume = 0 Then
                        'print the results
                        ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                        ws.Range("J" & 2 + j).Value = 0
                        ws.Range("K" & 2 + j).Value = "%" & 0
                        ws.Range("L" & 2 + j).Value = 0
    
                'If the cell immediately following a row is the same ticker...
                Else
                    'Find First non zero starting value
                    If ws.Cells(Start_Value, 3) = 0 Then
                        For find_value = Start_Value To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                    Start_Value = find_value
                                    Exit For
                            End If
                        Next find_value
                    End If
                    
                    'Calculate Change
                    Yearly_Change = (ws.Cells(i, 6) - ws.Cells(Start_Value, 3))
                    Percent_Change = Round((Yearly_Change / ws.Cells(Start_Value, 3) * 100), 2)
                    
                    'Start of the next stock ticker
                    Start_Value = i + 1
                    
                    'print the results to a separate worksheet
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = Yearly_Change
                    ws.Range("J" & 2 + j).NumberFormat = "0.00#"
                    ws.Range("K" & 2 + j).Value = "%" & Percent_Change
                    ws.Range("L" & 2 + j).Value = Total_Stock_Volume
                    
                    'colors positives green and negatives red
                    Select Case Yearly_Change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                End If
                   
                'reset variables for the new stock ticker
                Yearly_Change = 0
                Total_Stock_Volume = 0
                j = j + 1
            'If ticker is still the same add results
            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
           End If
                            
        Next i
    Next ws
End Sub
