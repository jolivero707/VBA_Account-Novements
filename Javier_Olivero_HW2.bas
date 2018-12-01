Attribute VB_Name = "Module1"
Option Explicit


Sub Ticker()

Dim WS_Count As Integer

Dim k As Integer

Dim j As Integer

Dim Ticker_open As Double

Dim Ticker_close As Double

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.

WS_Count = ActiveWorkbook.Worksheets.Count

' WS_Count = 1

k = 0

         ' Begin the loop.
For j = 1 To WS_Count

Sheets(j).Select

Dim i As Long

Dim Ticker As String

Dim TickerPosition As Long

Dim TotalVolume As Double

Dim LastRow As Long

Dim PercentChange As Double

' Determine the Last Row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

TotalVolume = 0

TickerPosition = 1

Ticker = ""

For i = 2 To LastRow

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
    Ticker = Cells(i, 1).Value
    
    TickerPosition = TickerPosition + 1
    
    TotalVolume = TotalVolume + Cells(i, 7).Value
    
    Range("L" & TickerPosition).Value = TotalVolume
    
    Ticker_close = Cells(i, 6).Value
    
'    Range("L" & TickerPosition).Value = Ticker_close
    
'   Range("K" & TickerPosition).Value = Ticker_open

    Range("J" & TickerPosition).Value = Ticker_close - Ticker_open
    
    If Ticker_open = 0 Then
    
         PercentChange = 0
    
    Else
         
         PercentChange = (Ticker_close - Ticker_open) / Ticker_open
    
    End If
    
    
    Range("K" & TickerPosition).Value = PercentChange
    
    Range("K" & TickerPosition).NumberFormat = "0.00%"
    
    If PercentChange >= 0 Then
    
       ' Set the Cell Colors to Green
        Range("J" & TickerPosition).Interior.ColorIndex = 4
    Else
        Range("J" & TickerPosition).Interior.ColorIndex = 3
        
    End If
    
    TotalVolume = 0
    
    k = 0
    
    Else
     
    TotalVolume = TotalVolume + Cells(i, 7).Value
    
        If k = 0 Then
    
            Ticker_open = Cells(i, 3).Value
    
            k = 1
    
        End If
    
    End If
    
    Range("I" & TickerPosition).Value = Ticker


Next i

Dim LastRow2 As Integer
Dim n As Integer
Dim MaxPercentage As Double
Dim MinPercentage As Double
Dim MaxVolume As Double
Dim Percentage As Double
Dim Volume As Double
Dim Max As Integer
Dim Min As Integer
Dim MaxV As Integer

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Volume Traded"

' Determine the Next Last Row
LastRow2 = Cells(Rows.Count, 9).End(xlUp).Row

MaxPercentage = Range("K2").Value
MinPercentage = Range("K2").Value
MaxVolume = Range("L2").Value

For n = 2 To LastRow2

Percentage = Range("K" & n).Value

If Percentage >= MaxPercentage Then
    MaxPercentage = Percentage
    Max = n
ElseIf Percentage < MinPercentage Then
    MinPercentage = Percentage
    Min = n
End If

Volume = Range("L" & n).Value

If Volume >= MaxVolume Then
    MaxVolume = Volume
    MaxV = n
End If

Next n


Range("P2").Value = Range("I" & Max).Value
Range("P3").Value = Range("I" & Min).Value
Range("P4").Value = Range("I" & MaxV).Value
Range("Q2").Value = MaxPercentage
Range("Q3").Value = MinPercentage
Range("Q4").Value = MaxVolume
Range("Q2").NumberFormat = "0.00%"
Range("Q3").NumberFormat = "0.00%"

Sheets(j).UsedRange.Columns.AutoFit

Next j

End Sub




