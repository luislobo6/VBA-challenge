Sub resumeStocks()

    'Declare variables
    Dim ticker, nextTicker As String
    Dim openValue, closeValue, percentChange, yearlyChange As Double
    Dim actualRow, nextRow, lastRow, resumeRow As Long
    Dim totalVolume As LongLong
    Dim isFirst As Boolean
        
    'Init variables
    resumeRow = 2
    actualRow = 2
    nextRow = actualRow + 1
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    isFirst = True
    openValue = 0
    closeValue = 0
    yearlyChange = 0
    percentChange = 0
    totalVolume = 0
    ticker = Cells(actualRow, 1).Value
    nesxtTicker = Cells(nextRow, 1).Value
    
    'Write Headers
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly_Change"
    Cells(1, 12).Value = "Percent_Change"
    Cells(1, 13).Value = "Total_Stock_Volume"
    
    'Start the cycle to check from the second row (data begins) to end of rows
    For i = 2 To lastRow
        
        'Each cicle add the Volume of the ticker and save values
        totalVolume = totalVolume + CLngLng(Cells(actualRow, 7).Value)
        actualRow = i
        nextRow = i + 1
        
        'If is the first time save the initial stock value
        If isFirst Then
            openValue = CDbl(Cells(actualRow, 3).Value)
            isFirst = False
        End If
        
        'Check the next cell until it changes or reaches the end of the rows to save values
        If ticker <> Cells(nextRow, 1).Value Or actualRow = lastRow Then
            
            closeValue = CDbl(Cells(actualRow, 6).Value) 'Close Value of Stock
            Cells(resumeRow, 10).Value = ticker 'Write ticker
            
            'Save yearly change and format it
            yearlyChange = closeValue - openValue
            Cells(resumeRow, 11).Value = yearlyChange
            If yearlyChange < 0 Then
                Cells(resumeRow, 11).Interior.ColorIndex = 3
            Else
                Cells(resumeRow, 11).Interior.ColorIndex = 4
            End If
            
            'Check if open value is 0 for the percent change
            If openValue = 0 Then
                'percentChange = closeValue * 100
                percentChange = 0
            Else
                percentChange = yearlyChange / openValue
            End If
            
            Cells(resumeRow, 12).Value = percentChange 'write percent change
            Cells(resumeRow, 13).Value = totalVolume 'Write total volume stock
            
            'save values for next ticker and restart variables
            resumeRow = resumeRow + 1
            isFirst = True
            openValue = 0
            closeValue = 0
            percentChange = 0
            totalVolume = 0
            ticker = Cells(nextRow, 1).Value
            
        End If
        
    Next i 'End of cycle
    
    'CHALLENGES
    
    'Set column and row names
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'Create and init variables
    Dim lastResumeRow As Long
    lastResumeRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "J").End(xlUp).Row
    Dim greatIncrease, geatDecrease As Double
    Dim greatTotalVol As LongLong
    Dim tickerGreatIncrease, tickerGreatDecrease, tickerTotVol As String
    greatIncrease = 0
    greatDecrease = 0
    greatTotalVol = 0
    
    'Iterate on the resume table to get values
    For x = 2 To CLng(lastResumeRow)
        
        'If the value on the cell is greater save cell value
        If greatIncrease < CDbl(Cells(x, 12).Value) Then
            greatIncrease = CDbl(Cells(x, 12).Value)
            tickerGreatIncrease = Cells(x, 10).Value
        End If
        
        'If the value on the cell is minor save cell value
        If greatDecrease > CDbl(Cells(x, 12).Value) Then
            greatDecrease = CDbl(Cells(x, 12).Value)
            tickerGreatDecrease = Cells(x, 10).Value
        End If
        
        'If the value on the cell is greater save cell value
        If greatTotalVol < CLngLng(Cells(x, 13).Value) Then
            greatTotalVol = CLngLng(Cells(x, 13).Value)
            tickerTotVol = Cells(x, 10).Value
        End If
        
    Next x
    
    'Write values on cells
    Range("P2").Value = tickerGreatIncrease
    Range("Q2").Value = greatIncrease
    Range("P3").Value = tickerGreatDecrease
    Range("Q3").Value = greatDecrease
    Range("P4").Value = tickerTotVol
    Range("Q4").Value = greatTotalVol
    
    'Autoadjust column width and format percent
    ActiveSheet.Columns("J:Q").AutoFit
    Columns("L:L").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Range("Q2:Q3").Style = "Percent"
    Range("Q2:Q3").NumberFormat = "0.00%"
    
End Sub

Sub resetResume()

    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "J").End(xlUp).Row
    
    Range("J1:Q" + Trim(Str(lastRow))).Value = ""
    Range("J1:Q" + Trim(Str(lastRow))).Interior.Color = xlNone
    ActiveSheet.Range("J:Q").ColumnWidth = 12
    ActiveSheet.Range("J:Q").NumberFormat = "General"
    
End Sub

Sub test()

    Dim test As LongLong
    test = CLngLng("123456789123456789")
    MsgBox (test)
End Sub
