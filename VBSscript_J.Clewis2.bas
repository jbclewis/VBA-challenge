Attribute VB_Name = "Module1"
Sub StockMarket()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Value"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

    Dim Yearly_Change As Double
    Dim Start As Double
    Dim j As Double
    Dim Percent_Change As Double
    Dim Total_Stock As Double
    Dim Largest As Double
    
    Yearly_Change = 0
    Start = 2
    j = 2
    Percent_Change = 0
    Total_Stock = 0
    
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        For Row = 2 To RowCount
        
        Total_Stock = Total_Stock + ws.Cells(Row, 7).Value
        
        If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
        
            Yearly_Change = (ws.Cells(Row, 6) - ws.Cells(Start, 3))
            ws.Range("I" & j).Value = ws.Cells(Row, 1).Value
            ws.Range("J" & j).Value = Yearly_Change
            ws.Range("J" & j).NumberFormat = "0.00"
        
            Percent_Change = (Yearly_Change / ws.Cells(Start, 3))
            ws.Range("K" & j).Value = Percent_Change
            ws.Range("K" & j).NumberFormat = "0.00%"
        
            ws.Range("L" & j).Value = Total_Stock
        
        If (ws.Cells(j, 10) <= 0) Then
            ws.Cells(j, 10).Interior.ColorIndex = 3 'Red
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 4 'Green
        End If
        
        If (ws.Cells(j, 11) <= 0) Then
            ws.Cells(j, 11).Interior.ColorIndex = 3 'Red
        Else
            ws.Cells(j, 11).Interior.ColorIndex = 4 'Green
        End If
        
        Yearly_Change = 0
        j = j + 1
        Start = Row + 1
        Percet_Change = 0
        Total_Stock = 0
        End If
        
        If ws.Cells(Row, 11) > Max Then
            Max = ws.Cells(Row, 11)
            Tag = ws.Cells(Row, 9)
            Cells(2, 15).Value = "Greatest % Increase"
            Cells(2, 16).Value = Tag
            Cells(2, 17).Value = Max
            Cells(2, 17).NumberFormat = "0.00%"
        End If
        
        If ws.Cells(Row, 11) < Min Then
            Min = ws.Cells(Row, 11)
            Tag = ws.Cells(Row, 9)
            Cells(3, 15).Value = "Greatest % Decrease"
            Cells(3, 16).Value = Tag
            Cells(3, 17).Value = Min
            Cells(3, 17).NumberFormat = "0.00%"
        End If
        
        If ws.Cells(Row, 12) > Largest Then
            Largest = ws.Cells(Row, 12)
            Tag = ws.Cells(Row, 9)
            Cells(4, 15).Value = "Greatest Total Volume"
            Cells(4, 16).Value = Tag
            Cells(4, 17).Value = Largest
        End If
    
        Next Row

Next ws
Columns("I:Q").AutoFit
End Sub




