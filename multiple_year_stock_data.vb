Sub FindTotalVolume():

' Declare variables
Dim i As Long
Dim j As Integer
' Declare summary row number
Dim isum As Integer

' Declare a variable to store total
Dim TotalVolume As Double
Dim ws As Worksheet


'loop through worksheets

For Each ws In Worksheets
  ws.Activate
  
  Dim yearBegin As Double
  Dim yearEnd As Double
  Dim percentChange As Double
  Dim percentGreatestIncrease As Double
  Dim percentGreatestDecrease As Double
  Dim GreatestTotalVolume As Double
  Dim GreatestTotalVolumeTicker As String
  
  
  ' Initalize total amount
  TotalVolume = 0
  yearlyChange = 0
  percentChange = 0

  ' Initialize summary row number
  isum = 2
  yearBegin = Cells(2, 3).Value

  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Stock Volume"
  percentGreatestIncrease = 0
  percentGreatestDecrease = 0
  GreatestTotalVolume = 0
  GreatestTotalVolumeTicker = " "
  percentGreatestIncreaseTicker = " "
  percentGreatestDecreaseTicker = " "

  ' Loop throught rows and columns
  For i = 2 To (Cells(Rows.Count, 1).End(xlUp).Row)
     TotalVolume = TotalVolume + Cells(i, 7).Value
     If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        Cells(isum, 9).Value = Cells(i, 1).Value
        yearEnd = Cells(i, 6).Value
        yearlyChange = yearEnd - yearBegin
        Cells(isum, 10).Value = yearlyChange
        If yearlyChange >= 0 Then
           Cells(isum, 10).Interior.ColorIndex = 4
        Else
           Cells(isum, 10).Interior.ColorIndex = 3
        End If
        
        If yearBegin <> 0 Then
           percentChange = (yearlyChange / yearBegin) * 100
        Else
           percentChange = 0
        End If
        
        Cells(isum, 11).Value = percentChange
        Cells(isum, 12).Value = TotalVolume
        
        
        If TotalVolume > GreatestTotalVolume Then
           GreatestTotalVolume = TotalVolume
           GreatestTotalVolumeTicker = Cells(i, 1).Value
        End If
        
        If percentChange > percentGreatestIncrease Then
           percentGreatestIncrease = percentChange
           percentGreatestIncreaseTicker = Cells(i, 1).Value
        End If
        
        If percentChange < percentGreatestDecrease Then
           percentGreatestDecrease = percentChange
           percentGreatestDecreaseTicker = Cells(i, 1).Value
        End If
        
        yearBegin = Cells(i + 1, 3).Value
        yearEnd = 0
        yearlyChange = 0
        isum = isum + 1
        TotalVolume = 0
     End If
  Next i
  
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

  
Cells(2, 15).Value = "Greatest % Increase"
Cells(2, 16).Value = percentGreatestIncreaseTicker
Cells(2, 17).Value = percentGreatestIncrease

Cells(3, 15).Value = "Greatest % Decrease"
Cells(3, 16).Value = percentGreatestDecreaseTicker
Cells(3, 17).Value = percentGreatestDecrease
  
Cells(4, 15).Value = "Greatest Total Volume"
Cells(4, 16).Value = GreatestTotalVolumeTicker
Cells(4, 17).Value = GreatestTotalVolume

Next ws

End Sub

