Attribute VB_Name = "Module1"
Sub StocksMY()

' Create a script that loops through all the stocks for one year and outputs the following information:
' 1. The ticker symbol
' 2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
' 3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
' 4. The total stock volume of the stock.
' Next, Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
' Lastly, Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

    'Declare variables
    Dim TotalVolume As LongLong
    Dim LastRow As LongLong
    Dim SummaryRowNo As Integer
    Dim YearOpenPrice As Double
    Dim YearClosePrice As Double
    Dim ws As Worksheet
    
    
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    'Autofit columns to text
    ws.Columns("A:Q").AutoFit
    
    'Reset variables for each worksheet
    TotalVolume = 0
    PriceChange = 0
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    SummaryRowNo = 2
    YearOpenPrice = Cells(2, 3).Value
    
    'Insert column headings
    Cells(1, 9).Value = "Ticker Symbol"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Create nested For Loop to go through data row by row
    
    For i = 2 To LastRow
        
        TotalVolume = TotalVolume + Cells(i, 7).Value 'Add stock volume line by line
        
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then 'Break if there is a change in ticker symbol
    
            Cells(SummaryRowNo, 9).Value = Cells(i, 1).Value 'Enter ticker symbol on summary line
            Cells(SummaryRowNo, 10).Value = Cells(i, 6).Value - YearOpenPrice 'Enter yearly change on summary line
                Range("J" & SummaryRowNo).NumberFormat = "#.00" 'Formats column to 2 dps
                'Add cell color if yearly change is positive/negative
                If Cells(SummaryRowNo, 10).Value >= 0 Then
                Cells(SummaryRowNo, 10).Interior.Color = vbGreen
                Else: Cells(SummaryRowNo, 10).Interior.Color = vbRed
                End If
            Cells(SummaryRowNo, 11).Value = FormatPercent(Cells(SummaryRowNo, 10).Value / YearOpenPrice)  'Calc and enter percent change
            Cells(SummaryRowNo, 12).Value = TotalVolume 'Enter Total Stock Volume for that stock
            TotalVolume = 0 'Reset to 0
            SummaryRowNo = SummaryRowNo + 1 'Move down one row of summary table
            YearOpenPrice = Cells(i + 1, 3).Value 'Set open price for next ticker/stock
        
        End If
        
    Next i
    
    'Add functionality to your script to return the stock with the "Greatest % increase",
    ' "Greatest % decrease", and "Greatest total volume".

    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVol As LongLong
    Dim GreatestIncrTick As String
    Dim GreatestDecrTick As String
    Dim GreatestVolTick As String
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"

    GreatestIncrease = Cells(2, 11).Value
    GreatestDecrease = Cells(2, 11).Value
    GreatestTotalVol = Cells(2, 12).Value
    
    For i = 2 To (SummaryRowNo - 1)
        'Calculate greatest increase
        If Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncrease = Cells(i, 11).Value
            GreatestIncrTick = Cells(i, 9).Value
        End If
        
        'Calculate greatest decrease
        If Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecrease = Cells(i, 11).Value
            GreatestDecrTick = Cells(i, 9).Value
        End If
        
        'Calculate greatest total volume
        If Cells(i, 12).Value > GreatestTotalVol Then
            GreatestTotalVol = Cells(i, 12).Value
            GreatestVolTick = Cells(i, 9).Value
        End If
    
    Next i
        
    'Enter values in spreadsheet
    Cells(2, 17).Value = FormatPercent(GreatestIncrease)
    Cells(2, 16).Value = GreatestIncrTick
    Cells(3, 17).Value = FormatPercent(GreatestDecrease)
    Cells(3, 16).Value = GreatestDecrTick
    Cells(4, 17).Value = GreatestTotalVol
    Cells(4, 16).Value = GreatestVolTick

Next ws

End Sub



