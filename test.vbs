Attribute VB_Name = "Module1"
Sub StockPractice()

Dim i, j As Long
Dim ws As Worksheet
Dim TotalVol As Double
Dim TotalVolA As Double
Dim YearChange As Double
Dim YearlyOpen As Double
Dim YearlyClose As Double
Dim LastRow As Long
Dim LastColumn As Double
Dim PercentChange As Double
Dim GreeatestPerInc As Double
Dim GreatestPerIncTicker As String
Dim GreatestPerDec As Double
Dim GreatestPerDecTicker As String
Dim GreatestTotVol As Double
Dim GreatestVolTicker As String
Dim MyRange As Range

' Define Headers and Greatest Change Values for Summary Tables

For Each ws In ThisWorkbook.Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Challenge Ticker"
        ws.Range("Q1").Value = "Challenge Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Cells.EntireColumn.AutoFit
       
' Define Starting Values

       TotalVol = 0
       YearlyOpen = ws.Cells(2, 3).Value
       

' Loop Function For Summary Table

       For i = 2 To ws.Range("A1").CurrentRegion.End(xlDown).Row
       LastRow = ws.Range("I1").CurrentRegion.Rows.Count + 1
       
        
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            TotalVol = TotalVol + ws.Cells(i, 7).Value
  
' Identifying Greatest Volume Change for Challenge
        Else
            ws.Range("I" & LastRow).Value = ws.Cells(i, 1).Value
            TotalVolA = ws.Cells(i, 7).Value
            TotalVol = TotalVol + TotalVolA
            ws.Range("L" & LastRow).Value = TotalVol
          
            YearlyClose = ws.Cells(i, 6).Value
            YearlyChange = YearlyClose - YearlyOpen
            ws.Range("J" & LastRow).Value = YearlyChange
            If ws.Range("J" & LastRow).Value >= 0 Then
' Conditional Indicator in Increase/Decrease Value
                ws.Range("J" & LastRow).Interior.ColorIndex = 4
            Else
                ws.Range("J" & LastRow).Interior.ColorIndex = 3
            End If
' Percentage Change for Summary Table
            If YearlyOpen = 0 Then
                PercentChange = 0
            Else
            PercentChange = (YearlyChange / YearlyOpen)
            End If
            
            ws.Range("K" & LastRow).Value = PercentChange
            ws.Range("K" & LastRow).NumberFormat = "0.00%"
' Identifying Greatest Percent Increase/Decrease for Challenge Table

    ' Reset Volume Total and Yearly Open
            TotalVol = 0
            YearlyOpen = ws.Cells(i + 1, 3).Value
        End If
             
        Next i
' Assigning Greatest Volume and Percent Changes to Challenge Summary Table
       GreatestPerInc = 0
       GreatestPerDec = 0
       GreatestTotVol = 0
       
       For j = 2 To LastRow
       
       If ws.Cells(j, 11).Value > GreatestPerInc Then
       GreatestPerInc = ws.Cells(j, 11).Value
       ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
       ws.Cells(2, 17).Value = GreatestPerInc
       ws.Cells(2, 17).NumberFormat = "0.00%"
       End If
       If ws.Cells(j, 11).Value < GreatestPerDec Then
       GreatestPerDec = ws.Cells(j, 11).Value
       ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
       ws.Cells(3, 17).Value = GreatestPerDec
       ws.Cells(3, 17).NumberFormat = "0.00%"
       End If
       If ws.Cells(j, 12).Value > GreatestTotVol Then
       GreatestTotVol = ws.Cells(j, 12).Value
       ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
       ws.Cells(4, 17).Value = GreatestTotVol
       End If
       Next j

        
    Next ws

End Sub



