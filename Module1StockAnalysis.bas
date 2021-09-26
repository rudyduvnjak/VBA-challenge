Attribute VB_Name = "Module1"

'Declare StockAnalysis Subroutine
Sub StockAnalysis()

'Define Variables

'Counts Numb
Dim CountLoops As Long
Dim Volume As Long
Dim PriceChange As Double
Dim PercentageChange As Double
Dim TotalVolume As Double
Dim PresentVolume As Long

Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim Ticker As String
Dim i As Long
Dim LastRow As Long
Dim j As Integer
Dim LineCount2 As Integer
LineCount2 = 2

'Set sht = ActiveWorksheet

'LastRow = sht.Worksheets(j).Cells(Rows.Count, 1).End(xlUp).Row
'Define variable that will keep number of tickers counted to move down the list
Dim CountTick As Integer
Dim column1 As Integer
column1 = 1
CountTick = 1
Dim Column7 As Integer
Column7 = 7

 
LastRow = ActiveSheet.UsedRange.Rows.Count
 
MsgBox LastRow

Range("K1") = "Ticker Symbol"
Range("L1") = "Yearly Change"
Range("M1") = "Percent Change"
Range("N1") = "Total Volume"
TotalVolume = 0


 


'Initial OpeningPrice entered by brute force, later values will be recorded inside loop
OpeningPrice = Range("C2").Value

For i = 2 To LastRow
    CountLoops = 1
    PresentVolume = Cells(i, Column7)
    'Sum the volume for current Ticker
    TotalVolume = (TotalVolume + PresentVolume)

    
    
    
    
    'Searches for when ticker symbol changes
    If Cells(i + 1, column1).Value <> Cells(i, column1) Then
    
    'Message Box the value of the current cell and value of the next cell
    'MsgBox (Cells(i, column1).Value & " and then " & Cells(i + 1, column1).Value)
    'Counts Ticker Number in the order to determine future write positions for all calculations
    CountTick = CountTick + 1
    Cells(CountTick, 11).Value = Cells(i, column1).Value
    CountLoops = (i + 1)
    'ClosingPrice value is determined by loopcount in Column6
    ClosingPrice = Cells(i, 6).Value
    PriceChange = ClosingPrice - OpeningPrice
    Cells(CountTick, 12).Value = PriceChange
    If PriceChange = 0 Or OpeningPrice = 0 Then
    PercentageChange = 0
    Else
    PercentageChange = ((PriceChange / OpeningPrice) * 100)
    End If
    PercentageChange = Round(PercentageChange, 2)
    If PercentageChange < 0 And PriceChange < 0 Then
    Cells(CountTick, 12).Interior.ColorIndex = 3
    Cells(CountTick, 13).Interior.ColorIndex = 3
    Else
    Cells(CountTick, 12).Interior.ColorIndex = 4
    Cells(CountTick, 13).Interior.ColorIndex = 4
    End If
    
    
    
    Cells(CountTick, 13).Value = (PercentageChange)
    Cells(CountTick, 14).Value = TotalVolume
    'If PriceChange.Value < 0 Then
    
    
    OpeningPrice = Cells(i + 1, 3).Value
    TotalVolume = 0
    
    End If
    
  Next i

End Sub


