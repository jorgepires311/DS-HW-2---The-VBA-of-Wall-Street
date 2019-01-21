Attribute VB_Name = "Module1"
Option Explicit

Sub analyzeStockMarketInit()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set up result columns
createResultHeadings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' calculate stock volume per ticker symbol
totalStockVolume
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' calculate stock change
yearlyStockChange
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' find greatest increase
findGreatestIncrease
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' find greatest decrease
findGreatestDecrease
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' convert formulas to values on current sheet
ConvertFormulasToValuesInActiveWorksheet
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' find greatest volume
findGreatestVolume
End Sub

Sub createResultHeadings()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim ticker As String
Dim currentRowTicker As String
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
Dim rowCount As Long
Dim resultRow As Long
Dim i As Long
Dim firstTickerRow As Long
Dim lastTickerRow As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = Cells(2, 1).End(xlDown).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set ticker
ticker = Cells(2, 1).Value
resultRow = 2
firstTickerRow = 2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set ticker
    currentRowTicker = Range("A" & i).Value
    If ticker <> currentRowTicker Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results for current ticker
        lastTickerRow = i - 1
        Range("I" & resultRow).Value = ticker
        Range("L" & resultRow).Formula = "=Sum(G" & firstTickerRow & ":G" & lastTickerRow & ")"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set new ticker to be tracked
        ticker = Range("A" & i).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' increment result row
        firstTickerRow = i
        resultRow = resultRow + 1
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
        If yearOpen <> 0 Then
            yearPrecentChange = yearChange / yearOpen
        Else
            yearPrecentChange = 0
        End If
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

Sub findGreatestIncrease()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim greatestIncreaseTicker As String
Dim greatestIncreasePercent As Double
Dim rowCount As Long
Dim i As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set initial values for variables
greatestIncreaseTicker = Cells(2, 9).Value

greatestIncreasePercent = Cells(2, 11).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = Cells(2, 9).End(xlDown).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
    If Cells(i, 11).Value > greatestIncreasePercent Then
        greatestIncreaseTicker = Cells(i, 9).Value
        greatestIncreasePercent = Cells(i, 11).Value
    End If
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results
Range("O2").Value = greatestIncreaseTicker
Range("P2").Value = greatestIncreasePercent
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' format results
Range("P2").NumberFormat = "0.00%"
End Sub

Sub findGreatestDecrease()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim greatestDecreaseTicker As String
Dim greatestDecreasePercent As Double
Dim rowCount As Long
Dim i As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set initial values for variables
greatestDecreaseTicker = Cells(2, 9).Value
greatestDecreasePercent = Cells(2, 11).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = Cells(2, 9).End(xlDown).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
    If Cells(i, 11).Value < greatestDecreasePercent Then
        greatestDecreaseTicker = Cells(i, 9).Value
        greatestDecreasePercent = Cells(i, 11).Value
    End If
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results
Range("O3").Value = greatestDecreaseTicker
Range("P3").Value = greatestDecreasePercent
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' format results
Range("P3").NumberFormat = "0.00%"
End Sub

Sub findGreatestVolume()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' create variables
Dim greatestVolumeTicker As String
Dim greatestVolumeValue As Double
Dim rowCount As Long
Dim i As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' set initial values for variables
greatestVolumeTicker = Cells(2, 9).Value
greatestVolumeValue = Cells(2, 12).Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' count total amount of rows
rowCount = Cells(2, 9).End(xlDown).Row
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' iterate through each row
For i = 2 To rowCount
    If Cells(i, 12).Value > greatestVolumeValue Then
        greatestVolumeTicker = Cells(i, 9).Value
        greatestVolumeValue = Cells(i, 12).Value
    End If
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' output results
Range("O4").Value = greatestVolumeTicker
Range("P4").Value = greatestVolumeValue
End Sub

Sub ConvertFormulasToValuesInActiveWorksheet()
Dim rng As Range
    For Each rng In ActiveSheet.UsedRange
        If rng.HasFormula Then
            rng.Formula = rng.Value
        End If
    Next rng
End Sub

