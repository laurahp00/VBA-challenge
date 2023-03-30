Attribute VB_Name = "Module1"
Sub VBA_challenge()

For Each ws In Worksheets

Dim WorksheetName As String
    WorksheetName = ws.Name
Dim ticker As Long
    ticker = 2
Dim percentChange As Double
Dim increase As Double
Dim decrease As Double
Dim vol As Double
Dim i As Long
Dim j As Long
    j = 2

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock vol"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total vol"


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To LastRow
    ' if next column =/= previous column, print next column and its
    ' complimentary values
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ' ticker name
        ws.Cells(ticker, 9).Value = ws.Cells(i, 1).Value
        
        ' yearly change
        ws.Cells(ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
        
        ' total stock volume
        ws.Cells(ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        
        ' percent change
        If ws.Cells(j, 3).Value <> 0 Then
            percentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
            ws.Cells(ticker, 11).Value = Format(percentChange, "Percent")
        Else
            ws.Cells(ticker, 11).Value = Format(0, "Percent")
        End If
        
        ' color-coding the yearly change
        If ws.Cells(ticker, 10).Value < 0 Then
            ws.Cells(ticker, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(ticker, 10).Interior.ColorIndex = 4
        End If
        
        ticker = ticker + 1
        j = i + 1
                
    End If
    
Next i

increase = ws.Cells(2, 11).Value ' first number in this line will begin the comparison
decrease = ws.Cells(2, 11).Value ' first number in this line will begin the comparison
vol = ws.Cells(2, 12).Value      ' first number in this line will begin the comparison

For i = 2 To LastRow2
    ' if the next cell is > than the prior, then that is the greater value
    If ws.Cells(i, 12).Value > vol Then
        vol = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    Else
        ' if not, the original value remains
        vol = vol
    End If

    'repeat logic with increase and decrease
    If ws.Cells(i, 11).Value > increase Then
        increase = ws.Cells(i, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    Else
        increase = increase
    End If

    If ws.Cells(i, 11).Value < decrease Then
        decrease = ws.Cells(i, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    Else
        decrease = decrease
        
    End If
            
Next i

' print the final values collected in loop
ws.Cells(2, 17).Value = Format(increase, "Percent")
ws.Cells(3, 17).Value = Format(decrease, "Percent")
ws.Cells(4, 17).Value = Format(vol, "Scientific")
Worksheets(WorksheetName).Columns("A:Z").AutoFit


Next ws
    
End Sub

