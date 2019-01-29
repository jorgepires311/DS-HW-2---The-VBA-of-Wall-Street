Sub analyzeStockMarketInit()
Dim ws As Worksheet
For Each ws In Worksheets
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set up result columns
Call createResultHeadings(ws)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' calculate stock volume per ticker symbol
Call totalStockVolume(ws)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' calculate stock change
Call yearlyStockChange(ws)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' find greatest increase
Call findGreatestIncrease(ws)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' find greatest decrease
Call findGreatestDecrease(ws)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' find greatest volume
Call findGreatestVolume(ws)
Next ws
End Sub

Sub createResultHeadings(sheet As Worksheet)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim ticker As String
Dim currentRowTicker As String
Dim rowCount As Long
Dim resultRow As Long
Dim i As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' prepare results columns
sheet.Range("I1").Value = "Ticker"
sheet.Range("J1").Value = "Yearly Change"
sheet.Range("K1").Value = "Percent Change"
sheet.Range("L1").Value = "Total Stock Volume"
sheet.Range("O1").Value = "Ticker"
sheet.Range("P1").Value = "Value"
sheet.Range("N2").Value = "Greatest % Increase"
sheet.Range("N3").Value = "Greatest % Decrease"
sheet.Range("N4").Value = "Greatest Total Value"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = sheet.Cells(Rows.Count, 1).End(xlUp).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set ticker
ticker = sheet.Cells(2, 1).Value
resultRow = 2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set ticker
    currentRowTicker = sheet.Range("A" & i).Value
    If ticker <> currentRowTicker Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results for current ticker
        sheet.Range("I" & resultRow).Value = ticker
        ticker = sheet.Range("A" & i).Value
        resultRow = resultRow + 1
    End If
Next i
    
End Sub

Sub totalStockVolume(sheet As Worksheet)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim ticker As String
Dim currentRowTicker As String
Dim rowCount As Long
Dim stockVolume As Double
Dim resultRow As Long
Dim i As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = sheet.Cells(Rows.Count, 1).End(xlUp).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set ticker
ticker = sheet.Cells(2, 1).Value
resultRow = 2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set ticker
    currentRowTicker = sheet.Range("A" & i).Value
    If ticker <> currentRowTicker Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results for current ticker
        sheet.Range("I" & resultRow).Value = ticker
        sheet.Range("L" & resultRow).Value = stockVolume
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set new ticker to be tracked
        ticker = sheet.Range("A" & i).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' increment result row
        resultRow = resultRow + 1
        stockVolume = 0
    End If
    stockVolume = stockVolume + sheet.Range("G" & i)
Next i
End Sub

Sub yearlyStockChange(sheet As Worksheet)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim ticker As String
Dim currentRowTicker As String
Dim stockVolume As Long
Dim rowCount As Long
Dim resultRow As Long
Dim i As Long
Dim lastRow As Long
Dim yearOpen As Double
Dim yearClose As Double
Dim yearChange As Double
Dim yearPrecentChange As Double
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = sheet.Cells(Rows.Count, 1).End(xlUp).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set initial values (ticker, year opening price and the first result row)
ticker = sheet.Cells(2, 1).Value
yearOpen = sheet.Cells(2, 3).Value
resultRow = 2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set current ticker
    currentRowTicker = sheet.Range("A" & i).Value
    If ticker <> currentRowTicker Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' calculate yearly change amount and percentage
        lastRow = i - 1 'go back up 1 row
        yearClose = sheet.Range("F" & lastRow).Value
        yearChange = yearClose - yearOpen
        If yearOpen <> 0 Then
            yearPrecentChange = yearChange / yearOpen
        Else
            yearPrecentChange = 0
        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results for current ticker
        sheet.Range("J" & resultRow).Value = yearChange
        sheet.Range("K" & resultRow).Value = yearPrecentChange
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' format results for current ticker
        If sheet.Range("J" & resultRow).Value > 0 Then
            sheet.Range("J" & resultRow).Interior.ColorIndex = 4 'green if positive
        Else
            sheet.Range("J" & resultRow).Interior.ColorIndex = 3 'red if negative
        End If
        sheet.Range("K" & resultRow).NumberFormat = "0.00%"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set next ticker
        ticker = sheet.Range("A" & i).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' collect next year open price
        yearOpen = sheet.Range("C" & i).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' increment result row
        resultRow = resultRow + 1
    End If
Next i
End Sub

Sub findGreatestIncrease(sheet As Worksheet)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim greatestIncreaseTicker As String
Dim greatestIncreasePercent As Double
Dim rowCount As Long
Dim i As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set initial values for variables
greatestIncreaseTicker = sheet.Cells(2, 9).Value

greatestIncreasePercent = sheet.Cells(2, 11).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = sheet.Cells(Rows.Count, 1).End(xlUp).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
    If sheet.Cells(i, 11).Value > greatestIncreasePercent Then
        greatestIncreaseTicker = sheet.Cells(i, 9).Value
        greatestIncreasePercent = sheet.Cells(i, 11).Value
    End If
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results
sheet.Range("O2").Value = greatestIncreaseTicker
sheet.Range("P2").Value = greatestIncreasePercent
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' format results
sheet.Range("P2").NumberFormat = "0.00%"
End Sub

Sub findGreatestDecrease(sheet As Worksheet)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim greatestDecreaseTicker As String
Dim greatestDecreasePercent As Double
Dim rowCount As Long
Dim i As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set initial values for variables
greatestDecreaseTicker = sheet.Cells(2, 9).Value
greatestDecreasePercent = sheet.Cells(2, 11).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = sheet.Cells(Rows.Count, 1).End(xlUp).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
    If sheet.Cells(i, 11).Value < greatestDecreasePercent Then
        greatestDecreaseTicker = sheet.Cells(i, 9).Value
        greatestDecreasePercent = sheet.Cells(i, 11).Value
    End If
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results
sheet.Range("O3").Value = greatestDecreaseTicker
sheet.Range("P3").Value = greatestDecreasePercent
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' format results
sheet.Range("P3").NumberFormat = "0.00%"
End Sub

Sub findGreatestVolume(sheet As Worksheet)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim greatestVolumeTicker As String
Dim greatestVolumeValue As Double
Dim rowCount As Long
Dim i As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set initial values for variables
greatestVolumeTicker = sheet.Cells(2, 9).Value
greatestVolumeValue = sheet.Cells(2, 12).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = sheet.Cells(Rows.Count, 1).End(xlUp).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
    If sheet.Cells(i, 12).Value > greatestVolumeValue Then
        greatestVolumeTicker = sheet.Cells(i, 9).Value
        greatestVolumeValue = sheet.Cells(i, 12).Value
    End If
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results
sheet.Range("O4").Value = greatestVolumeTicker
sheet.Range("P4").Value = greatestVolumeValue
End Sub

