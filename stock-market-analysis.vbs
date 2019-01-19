Sub analyzeStockMarketInit()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set up result columns
createResultHeadings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' calculate stock volume per ticker symbol
'totalStockVolume
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' calculate stock change
yearlyStockChange
End Sub

Sub createResultHeadings()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim ticker As String
Dim rowCount As Long
Dim resultRow As Long
Dim i As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' prepare results columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Value"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = Cells(2, 1).End(xlDown).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set ticker
ticker = Cells(2, 1).Value
resultRow = 2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set ticker
    currentRowTicker = Range("A" & i).Value
    If ticker <> currentRowTicker Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results for current ticker
        Range("I" & resultRow).Value = ticker
        ticker = Range("A" & i).Value
        resultRow = resultRow + 1
    End If
Next i
    
End Sub

Sub totalStockVolume()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim ticker As String
Dim currentRowTicker As String
Dim stockVolume As Long
Dim rowCount As Long
Dim resultRow As Long
Dim i As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = Cells(2, 1).End(xlDown).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set ticker
ticker = Cells(2, 1).Value
resultRow = 2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set ticker
    currentRowTicker = Range("A" & i).Value
    If ticker <> currentRowTicker Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results for current ticker
        Range("I" & resultRow).Value = ticker
        Range("L" & resultRow).Value = stockVolume
        ticker = Range("A" & i).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' increment result row
        stockVolume = 0
        resultRow = resultRow + 1
    Else
        stockVolume = stockVolume + Cells(i, 7)
    End If
Next i
End Sub

Sub yearlyStockChange()
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
rowCount = Cells(2, 1).End(xlDown).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set initial values (ticker, year opening price and the first result row)
ticker = Cells(2, 1).Value
yearOpen = Cells(2, 3).Value
resultRow = 2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set current ticker
    currentRowTicker = Range("A" & i).Value
    If ticker <> currentRowTicker Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' calculate yearly change amount and percentage
        lastRow = i - 1 'go back up 1 row
        yearClose = Range("F" & lastRow).Value
        yearChange = yearClose - yearOpen
        yearPrecentChange = yearChange / yearClose
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results for current ticker
        Range("J" & resultRow).Value = yearChange
        Range("K" & resultRow).Value = yearPrecentChange
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' format results for current ticker
        If Range("J" & resultRow).Value > 0 Then
            Range("J" & resultRow).Interior.ColorIndex = 4 'green if positive
        Else
            Range("J" & resultRow).Interior.ColorIndex = 3 'red if negative
        End If
        Range("K" & resultRow).NumberFormat = "0.00%"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set next ticker
        ticker = Range("A" & i).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' collect next year open price
        yearOpen = Range("C" & i).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' increment result row
        resultRow = resultRow + 1
    End If
Next i
End Sub


