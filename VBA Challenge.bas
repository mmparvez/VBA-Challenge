Attribute VB_Name = "Module1"
Sub ticker()

'Iterate through each sheet
For Each ws In Worksheets

'Declare all variables
    Dim LastRow As Long
    Dim i As Integer
    Dim high_value As Double
    Dim low_value As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim ttl_vol As Double
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Declare innitial value
    ttl_vol = 0
    i = 2
    r = 2
    
    'Find the last row number in each sheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Open price of the firstday of the year
    open_price = ws.Cells(r, 3).Value
    
    'Itterate through each row
    For r = 2 To LastRow
        'Condition to find new value in the first column
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            ws.Cells(i, 9).Value = ws.Cells(r, 1).Value
            
            'Adding last vol data to the total and writing Total vol dta in specifiq cell
            ttl_vol = (ttl_vol + ws.Cells(r, 7).Value)
            ws.Cells(i, 12).Value = ttl_vol
            
            'Innitialise variable for new ticker
            ttl_vol = 0
            
            
            'Calculate and write yearly chage data
            close_price = ws.Cells(r, 6).Value
            ws.Cells(i, 10).Value = (close_price - open_price)
                        
            'Conditional formatting yearly change column
            If ws.Cells(i, 10).Value <= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 4
            End If
            
            'Calculating and change format ofpercent change
            ws.Cells(i, 11).Value = (close_price - open_price) / open_price
            ws.Cells(i, 11).NumberFormat = "0.00%"
            
            open_price = ws.Cells(r + 1, 3).Value
            i = i + 1
        Else
            'Summing volume data when condition is false
            ttl_vol = (ttl_vol + ws.Cells(r, 7).Value)
        End If
    Next r
    
Next ws

End Sub

Sub bonus()

'Iterate through each sheet
For Each ws In Worksheets

'Declare all variables
    Dim LastRow As Integer
    Dim i As Integer
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_vol As Double
    Dim vol As Long
    Dim maxname As String
    
    'Header on top row and row lebel
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Find the last row number in each sheet
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Finding max value of percentage change column
    max_increase = WorksheetFunction.Max(ws.Range("K2", "K" & LastRow))
    
    'Finding respective ticker of maximum percentage change
    maxname = Application.WorksheetFunction.Index(ws.Range("I2", "I" & LastRow), Application.WorksheetFunction.Match(max_increase, ws.Range("K2", "K" & LastRow), 0))
    'Writing findings on specific cells
    ws.Cells(2, 16).Value = maxname
    ws.Cells(2, 17).Value = max_increase
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    'Finding min value of percentage change column
    max_decrease = WorksheetFunction.Min(ws.Range("K2", "K" & LastRow))
    
    'Finding respective ticker of minimum percentage change
    maxname = Application.WorksheetFunction.Index(ws.Range("I2", "I" & LastRow), Application.WorksheetFunction.Match(max_decrease, ws.Range("K2", "K" & LastRow), 0))
    'Writing findings on specific cells
    ws.Cells(2, 16).Value = maxname
    ws.Cells(3, 16).Value = maxname
    ws.Cells(3, 17).Value = max_decrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    'Finding max value of Total Stock Volume column
    max_vol = WorksheetFunction.Max(ws.Range("L2", "L" & LastRow))
    
    'Finding respective ticker of maximum volume
    maxname = Application.WorksheetFunction.Index(ws.Range("I2", "I" & LastRow), Application.WorksheetFunction.Match(max_vol, ws.Range("L2", "L" & LastRow), 0))
    'Writing findings on specific cells
    ws.Cells(4, 16).Value = maxname
    ws.Cells(4, 17).Value = max_vol
  
    
    
Next ws
    
    
End Sub
