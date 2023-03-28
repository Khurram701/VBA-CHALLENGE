Sub Uniquetickers()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickers As Range
    Dim Uniquetickers As Variant
    For Each ws In Worksheets
      lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
      Set tickers = ws.Range("A1:A" & lastRow)
      Uniquetickers = tickers.AdvancedFilter(Action:=xlFilterCopy, CopyToRange:=ws.Range("I1"), Unique:=True)
Next ws

End Sub



Sub CalculateYearlyandpercentChange()
    
    ' Define variables
    Dim ticker As String
    Dim lastRow As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim yearlyChange As Double
    Dim SummaryTableRowIndex As Integer
    Dim percentChange As Double
   
    For Each ws In Worksheets
    SummaryTableRowIndex = 1
    OpenPrice = 0
    ClosePrice = 0
    
    lastRow = ws.Range("A1").End(xlDown).Row
    
    For i = 2 To lastRow
        
        If ws.Cells(i, 1).Value <> ticker Then
            
            If OpenPrice <> 0 Then
                yearlyChange = ClosePrice - OpenPrice
                ws.Cells(SummaryTableRowIndex, 10).Value = yearlyChange
               
                
                percentChange = yearlyChange / OpenPrice
                ws.Cells(SummaryTableRowIndex, 11).Value = FormatPercent(percentChange, 11)
            End If
            
            ticker = ws.Cells(i, 1).Value
            OpenPrice = ws.Cells(i, 3).Value
            SummaryTableRowIndex = SummaryTableRowIndex + 1
            
        End If
        
        ' Set the close price for the current row
        ClosePrice = ws.Cells(i, 6).Value
        
    Next i
    
    ' Calculate the yearly change and add it to the summary table for the last ticker symbol
    yearlyChange = ClosePrice - OpenPrice
   percentChange = yearlyChange / OpenPrice
   
   ws.Cells(SummaryTableRowIndex, 10).Value = yearlyChange
   ws.Cells(SummaryTableRowIndex, 11).Value = FormatPercent(percentChange, 11)
   
   
   
Next ws
   
   
    
End Sub

Sub CalculateTotalVolume()
    Dim lastRow As Long
    Dim ticker As String
    Dim volume As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    
    For Each ws In Worksheets
    
    ' Find the last row of data in column A
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set the initial values for the ticker and total volume
    ticker = ws.Range("A2")
    totalVolume = ws.Range("G2")
    outputRow = 2
    
    ' Loop through each row of data
    For i = 3 To lastRow
        ' Check if this is the same ticker as the previous row
        If ws.Range("A" & i) = ticker Then
            ' If it is, add the volume to the total
            volume = ws.Range("G" & i)
            totalVolume = totalVolume + volume
        Else
            ' If it isn't, write the total volume for the previous ticker and reset the total volume for the new ticker
            ws.Range("L" & outputRow) = totalVolume
            outputRow = outputRow + 1
            ticker = ws.Range("A" & i)
            totalVolume = ws.Range("G" & i)
        End If
    Next i
    
    ' Write the total volume for the final ticker
    ws.Range("L" & outputRow) = totalVolume
    
    Next ws
    
End Sub

Sub FormatYearlyChange()
    Dim lastRow As Long
    Dim cell As Range
    
    For Each ws In Worksheets
    
    ' Find the last row of data in column A
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each cell in the yearly change column
    For Each cell In ws.Range("J2:J" & lastRow)
        ' Check if the value is positive
        If cell.Value > 0 Then
            ' If it is, set the cell's interior color to green
            cell.Interior.Color = RGB(0, 255, 0)
        ElseIf cell.Value < 0 Then
            ' If it isn't, set the cell's interior color to red
            cell.Interior.Color = RGB(255, 0, 0)
        End If
    Next cell
    
    Next ws
End Sub


Sub GreatestPercentIncreaseAcrossSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim tickerRange As Range
    Dim tickerCell As Range
    Dim maxPercentIncrease As Double
    Dim maxTicker As String
    Dim sheetNames As Variant
    Dim i As Integer
    
    sheetNames = Array("2018", "2019", "2020")
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = Worksheets(sheetNames(i))
        
        'Find last row
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        'Set ticker range
        Set tickerRange = ws.Range("I2:I" & lastRow)
        
        'Loop through each ticker
        For Each tickerCell In tickerRange
            'Check if % increase is greater than current max % increase
            If tickerCell.Offset(0, 2).Value > maxPercentIncrease Then
                'Set new max % increase and ticker
                maxPercentIncrease = tickerCell.Offset(0, 2).Value
                maxTicker = tickerCell.Value
            End If
        Next tickerCell
        
        'Set values in output cells
        ws.Range("P2").Value = maxTicker
        ws.Range("Q2").Value = Format(maxPercentIncrease, "0.00%")
        
        'Reset max values for next sheet
        maxPercentIncrease = 0
        maxTicker = ""
    Next i
End Sub

Sub CalculateGreatestPercentDecrease()
    
    ' Define variables
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim maxDecrease As Double
    Dim maxDecreaseTicker As String
    
    For Each ws In Worksheets
    
    ' Initialize variables
    maxDecrease = 0
    
    ' Loop through each row in the data
    For i = 2 To ws.Range("I" & Rows.Count).End(xlUp).Row
        
        ' Get values from the current row
        ticker = ws.Cells(i, "I").Value
        yearlyChange = ws.Cells(i, "J").Value
        percentChange = ws.Cells(i, "K").Value
        totalVolume = ws.Cells(i, "L").Value
        
        ' Check if the current percent change is the greatest decrease so far
        If percentChange < maxDecrease Then
            maxDecrease = percentChange
            maxDecreaseTicker = ticker
        End If
        
    Next i
    
    ' Print the result to the worksheet
    ws.Range("P3").Value = maxDecreaseTicker
    ws.Range("Q3").Value = maxDecrease
    
    Next ws
    
End Sub

Sub CalculateGreatestTotalVolume()
    
    ' Define variables
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim maxVolume As Double
    Dim maxVolumeTicker As String
    
    For Each ws In Worksheets
    
    ' Initialize variables
    maxVolume = 0
    
    ' Loop through each row in the data
    For i = 2 To ws.Range("I" & Rows.Count).End(xlUp).Row
        
        ' Get values from the current row
        ticker = ws.Cells(i, "I").Value
        yearlyChange = ws.Cells(i, "J").Value
        percentChange = ws.Cells(i, "K").Value
        totalVolume = ws.Cells(i, "L").Value
        
        ' Check if the current volume is the greatest so far
        If totalVolume > maxVolume Then
            maxVolume = totalVolume
            maxVolumeTicker = ticker
        End If
        
    Next i
    
    ' Print the result to the worksheet
    ws.Range("P4").Value = maxVolumeTicker
    ws.Range("Q4").Value = maxVolume
    
    Next ws
    
End Sub

Sub FormatPercentChange()
    Dim lastRow As Long
    Dim cell As Range
    
    For Each ws In Worksheets
    
    ' Find the last row of data in column A
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each cell in the yearly change column
    For Each cell In ws.Range("K2:K" & lastRow)
        ' Check if the value is positive
        If cell.Value > 0 Then
            ' If it is, set the cell's interior color to green
            cell.Interior.Color = RGB(0, 255, 0)
        ElseIf cell.Value < 0 Then
            ' If it isn't, set the cell's interior color to red
            cell.Interior.Color = RGB(255, 0, 0)
        End If
    Next cell
    
    Next ws
End Sub

