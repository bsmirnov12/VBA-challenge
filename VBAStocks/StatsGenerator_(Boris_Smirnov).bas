Attribute VB_Name = "Module1"
Public Const tickerNamePos = 1
Public Const openingPricePos = 3
Public Const closingPricePos = 6
Public Const totalVolumePos = 7
Public Const tickerStatsPos = 9
Public Const greatestStatsPos = 15
Public Const colorPositive = 4 'Green
Public Const colorNegative = 3 'Red


Sub StatsGenerator()
Attribute StatsGenerator.VB_ProcData.VB_Invoke_Func = "x\n14"
'Create ticker stats
'(c) Boris Smirnov
'UT-TOR-DATA-PT-01-2020-U-C / 02-VBA-Scripting Homework assignment

    For Each ws In Worksheets
        lastTickerRow = CreateTickerStats(ws)
        Call CreateGreatestStats(ws, lastTickerRow)
    Next
    
End Sub

Sub CreateGreatestStats(ws, lastTickerRow)
'Adds the greatest values stats onto given worksheet
'Parameters
'   ws (Inbput)
'       worksheet to work on
'   lastTickerRow (Input)
'       row  number of the last ticker on the sheet

    'Creating stats table row and column names
    rangePos = ColNum2Letter(greatestStatsPos + 1) + "1:" + ColNum2Letter(greatestStatsPos + 2) + "1"
    ws.Range(rangePos).Value = Array("Ticker", "Value")

    rangePos = ColNum2Letter(greatestStatsPos) + "2:" + ColNum2Letter(greatestStatsPos) + "4"
    ws.Range(rangePos).Value = Application.Transpose(Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume"))
    
    'Finding extremums
    Dim maxPcInc, maxPcDec As Double
    Dim maxTotal As LongLong
    Dim rowPcInc, rowPcDec, rowTotal As Long
    maxPcInc = 0#
    maxPcDec = 0#
    maxTotal = 0
    
    For i = 2 To lastTickerRow
        'Greatest % Increase
        If ws.Cells(i, tickerStatsPos + 2).Value > maxPcInc Then
            maxPcInc = ws.Cells(i, tickerStatsPos + 2).Value
            rowPcInc = i
        'Greatest % Decrease
        ElseIf ws.Cells(i, tickerStatsPos + 2).Value < maxPcDec Then
            maxPcDec = ws.Cells(i, tickerStatsPos + 2).Value
            rowPcDec = i
        End If
        
        'Greatest Total Volume
        If ws.Cells(i, tickerStatsPos + 3).Value > maxTotal Then
            maxTotal = ws.Cells(i, tickerStatsPos + 3).Value
            rowTotal = i
        End If
    Next i
    
    'Reporting extremums
    'Greatest % Increase
    ws.Cells(2, greatestStatsPos + 1).Value = ws.Cells(rowPcInc, tickerStatsPos).Value 'Ticker name
    ws.Cells(2, greatestStatsPos + 2).Value = maxPcInc
    ws.Cells(2, greatestStatsPos + 2).NumberFormat = "0.00%"
    'Greatest % Decrease
    ws.Cells(3, greatestStatsPos + 1).Value = ws.Cells(rowPcDec, tickerStatsPos).Value 'Ticker name
    ws.Cells(3, greatestStatsPos + 2).Value = maxPcDec
    ws.Cells(3, greatestStatsPos + 2).NumberFormat = "0.00%"
    'Greatest Total Volume
    ws.Cells(4, greatestStatsPos + 1).Value = ws.Cells(rowPcInc, tickerStatsPos).Value 'Ticker name
    ws.Cells(4, greatestStatsPos + 2).Value = maxTotal
    ws.Cells(4, greatestStatsPos + 2).NumberFormat = "###,###,###,###"
   
End Sub

Function CreateTickerStats(ws)
'The function creates ticker stats part on the sheet passed as input parameter and returns
' row  number of the last ticker
'Parameters:
'   ws (Input) - worksheet to scan

    Dim startRow, scanRow, statsRow As Long
    Dim tickerName As String
    Dim yearStart, yearEnd As Double
    
    'Initializing
    startRow = 2
    scanRow = startRow
    statsRow = 2
    
    'Create tickers stats header
    rangePos = ColNum2Letter(tickerStatsPos) + "1:" + ColNum2Letter(tickerStatsPos + 3) + "1"
    ws.Range(rangePos).Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
    'Scanning tickers one by one
    Do Until ws.Cells(scanRow, tickerNamePos).Value = ""
        Call ScanTicker(ws, scanRow, tickerName, yearStart, yearEnd)
        Call PrintTickerStats(ws, statsRow, tickerName, yearStart, yearEnd, startRow, scanRow)
        startRow = scanRow
    Loop
        
    CreateTickerStats = statsRow - 1
End Function

Sub PrintTickerStats(ws, ByRef rowNumber, tickerName, yearStart, yearEnd, startRow, endRow)
'The procedure prints a row of statistics for one ticker at specific position set by rowNumber parameter and tickerStatsPos constant
'Parameters:
'   ws (Input)
'       worksheet to print on
'   rowNumber (Input/Output)
'       the row where to print stats, then used as a return value to set the next row
'   tickerName, yearStart, yearEnd (Input)
'       ticker name, and corresponding year opening and closing prices
'   startRow, endRow (Input)
'       row numbers,  [start, end), of the data on the sheet for given ticker
'       used to calculate Total Stock Volume

    Dim priceDiff As Double
    priceDiff = yearEnd - yearStart
    
    ws.Cells(rowNumber, tickerStatsPos).Value = tickerName
    
    ws.Cells(rowNumber, tickerStatsPos + 1).Value = priceDiff
    If priceDiff < 0 Then
        ws.Cells(rowNumber, tickerStatsPos + 1).Interior.ColorIndex = colorNegative
    ElseIf priceDiff > 0 Then
        ws.Cells(rowNumber, tickerStatsPos + 1).Interior.ColorIndex = colorPositive
    End If
    
    If Not yearStart = 0 Then
        ws.Cells(rowNumber, tickerStatsPos + 2).Value = priceDiff / yearStart
        ws.Cells(rowNumber, tickerStatsPos + 2).NumberFormat = "0.00%"
    End If
    
    ws.Cells(rowNumber, tickerStatsPos + 3).Formula = "=SUM(" + ColNum2Letter(totalVolumePos) + CStr(startRow) + ":" + ColNum2Letter(totalVolumePos) + CStr(endRow - 1) + ")"
    ws.Cells(rowNumber, tickerStatsPos + 3).NumberFormat = "###,###,###,###"
    
    rowNumber = rowNumber + 1
End Sub

Sub ScanTicker(ws, ByRef rowNumber, ByRef tickerName, ByRef yearStart, ByRef yearEnd)
'The procedure scans current sheet row by row until ticker name changes
' and returns opening and closing prices for the ticker, the row number after ticker name changed,
' and ticker name
'Parameters:
'   ws (Input)
'       worksheet to scan
'   rowNumber (Input/Output)
'       as input parameter passes start row number on the current sheet from where scan begins
'       as output parameter returns row number where ticker name changed or the data ended
'       (the next row after the last with current ticker name)
'   tickerName (Output)
'       returns tickerName from the first column at the beginning of the scan
'   yearStart (Output)
'       returns value of the open price of the first day of the year
'   yearEnd (Output)
'       returns value of the closing price of the last day of the year
    
    tickerName = ""
    yearStart = 0#
    yearEnd = 0#
    If ws.Cells(rowNumber, tickerNamePos).Value = "" Then
        Exit Sub
    End If
    
    'Initializing
    tickerName = CStr(ws.Cells(rowNumber, tickerNamePos).Value)
    yearStart = CDbl(ws.Cells(rowNumber, openingPricePos).Value)
    
    'Scanning data until ticker name changes
    While CStr(ws.Cells(rowNumber, tickerNamePos).Value) = tickerName
        rowNumber = rowNumber + 1
    Wend
    
    'Getting year end price
    yearEnd = CDbl(ws.Cells(rowNumber - 1, closingPricePos).Value)
    
    'At this point rowNumber has the next row numer after the data pertaining tickerName
End Sub

Function ColNum2Letter(colNum)
'The function converts column number to letter
    If colNum <= 26 Then
        'My solution
        ColNum2Letter = String(1, Chr(Asc("A") + (colNum - 1)))
    Else
        'Borrowed from https://www.thespreadsheetguru.com/the-code-vault/vba-code-to-convert-column-number-to-letter-or-letter-to-number
        ColNum2Letter = Split(Cells(1, colNum).Address, "$")(1)
    End If
End Function

