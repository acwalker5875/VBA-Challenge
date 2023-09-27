Attribute VB_Name = "Module1"
Sub stocks()

Dim lastrow As Double
Dim i As Double
Dim ws As Worksheet
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim TotalVolume As Double
Dim stock1 As String
Dim stock2 As String

'Apply to each Worksheet

For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
       
       'Add Cell Titles
       
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Year Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
        
  'set counters for volume and group
    TotalVolume = 0
    StockGroup = 1
        
'Define last row
lastrow = Cells(ws.Rows.Count, 1).End(xlUp).Row

'Define first opening price
OpenPrice = Cells(2, 3).Value

For i = 2 To lastrow
    
    stock1 = Cells(i, 1).Value
    stock2 = Cells(i + 1, 1).Value
    
    If stock1 = stock2 Then
    
        'Count Stock Volume
        TotalVolume = TotalVolume + Cells(i, 7).Value
    
    Else
        'Define Final Values of Stock
        TotalVolume = TotalVolume + Cells(i, 7).Value
        ClosePrice = Cells(i, 6).Value
         YrChng = ClosePrice - OpenPrice
         PctChng = (ClosePrice / OpenPrice) - 1
        
         
        'apply name to ticker
        Cells(StockGroup + 1, 9).Value = stock1
        
        'place values in ticker
        Cells(StockGroup + 1, 10).Value = YrChng
      Cells(StockGroup + 1, 11).Value = PctChng
       Cells(StockGroup + 1, 12).Value = TotalVolume
    
    'Reset Values for Next Stock Value
    TotalVolume = 0
    OpenPrice = Cells(i + 1, 3).Value
    StockGroup = StockGroup + 1
    
    
    End If
    
    'set colors for positive/negative yearly changes
       If Cells(i, 10) > 0 And Len(Cells(i, 10)) > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
            ElseIf Cells(i, 10) <= 0 And Len(Cells(i, 10)) > 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    'set colors for positive/negative percent changes
    If Cells(i, 11) > 0 And Len(Cells(i, 11)) > 0 Then
            Cells(i, 11).Interior.ColorIndex = 4
            ElseIf Cells(i, 11) <= 0 And Len(Cells(i, 11)) > 0 Then
            Cells(i, 11).Interior.ColorIndex = 3
        End If
    
    Next i
    
    'Format Values
    Range("K2:K" & lastrow).NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"
    

    
     'find Summary Values
    Cells(2, 17).Value = WorksheetFunction.Max(Columns("k"))
    Cells(3, 17).Value = WorksheetFunction.Min(Columns("k"))
    Cells(4, 17).Value = WorksheetFunction.Max(Columns("l"))

    MaxIncrease = WorksheetFunction.Match(Cells(2, 17).Value, Range("K:K"), 0)
    MaxDecrease = WorksheetFunction.Match(Cells(3, 17).Value, Range("K:K"), 0)
    MaxTotal = WorksheetFunction.Match(Cells(4, 17).Value, Range("L:L"), 0)
    
    Cells(2, 16).Value = Cells(MaxIncrease, 9)
    Cells(3, 16).Value = Cells(MaxDecrease, 9)
    Cells(4, 16).Value = Cells(MaxTotal, 9)

Next

    
    

End Sub

