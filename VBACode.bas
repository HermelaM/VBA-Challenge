VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

 Sub StockData()
 
 'Loop through all worksheets
  For Each ws In Worksheets

        'Declare Variables and initial value
        Dim TickerName As String
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalStockVol As Double
        TotalStockVol = 0
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        Dim PreviousAmount As Long
        PreviousAmount = 2
         
        'Challenge Variables
        
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim LastRowValue As Long
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0

        'Column Headers
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Challenge row/column headers
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ' Find the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For I = 2 To LastRow

            ' Add To Ticker Total Volume with in same Ticker symbole
             TotalStockVol = TotalStockVol + ws.Cells(I, 7).Value
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then


                ' Set Ticker Name
                TickerName = ws.Cells(I, 1).Value
                
                 ' Print The Ticker
                ws.Range("I" & SummaryTableRow).Value = TickerName
                
                ' Set Yearly Change
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & I)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
                
                ' Print The Ticker Total Amount To The Summary Table
                ws.Range("L" & SummaryTableRow).Value = TotalStockVol
                ' Reset Ticker Total
                 TotalStockVol = 0

                ' Determine Percent Change
                If YearlyOpen = 0 Then
                   PercentChange = 0
                   
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                    
                End If
                
                ' Adding % symbol
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                ' Conditional Formatting Highlight Positive (Green) / Negative (Red)
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                ' Add One To The Summary Table Row
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = I + 1
                End If
            Next I
   
   'Challenge comparison for greatest increase/decrease and total

            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            ' Start Loop For Final Results
            For I = 2 To LastRow
                If ws.Range("K" & I).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & I).Value
                    ws.Range("P2").Value = ws.Range("I" & I).Value
                End If

                If ws.Range("K" & I).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & I).Value
                    ws.Range("P3").Value = ws.Range("I" & I).Value
                End If

                If ws.Range("L" & I).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & I).Value
                    ws.Range("P4").Value = ws.Range("I" & I).Value
                End If

            Next I
            
        'Adding % symbol
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
 
    Next ws

End Sub

