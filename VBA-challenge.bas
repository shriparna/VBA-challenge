Attribute VB_Name = "Module1"
Sub StockTicker():

    ' Define variables
    Dim TotalRowCount As Double
    Dim PrintRow As Double
    Dim StartTime As Double
    Dim EndTime As Double
    
    Dim TickerChanged As Boolean
    Dim OpenValue As Double
    Dim CloseValue As Double
    
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalStockVolume As Double
    
    ' Define variables for the second challenge
    Dim GPI As Double
    Dim GPD As Double
    Dim TSV As Double
    Dim TickerGPI As String
    Dim TickerGDP As String
    Dim TickerTSV As String
        
    StartTime = Timer
    ' Search across all worksheets
    For Each wsheet In Worksheets
    
        ' Reset the column by deleting the existing calculation
        wsheet.Range("I:Q").Delete
        MsgBox ("Press OK to Start " + wsheet.Name)
        
        ' Initialize the variables
        ' Get the TotalRowCount for the current worksheet
        TotalRowCount = wsheet.Cells(Rows.Count, 1).End(xlUp).Row
        PrintRow = 1
        TickerChanged = True
        
        ' First print the column headers in the current worksheet
        wsheet.Range("I1").Value = "Ticker"
        wsheet.Range("J1").Value = "Yearly Change"
        wsheet.Range("K1").Value = "Percent Change"
        wsheet.Range("L1").Value = "Total Stock Volume"
        
        ' Print the titles for the second challenge
        wsheet.Range("P1").Value = "Ticker"
        wsheet.Range("Q1").Value = "Value"
        wsheet.Range("O2").Value = "Greatest % Increase"
        wsheet.Range("O3").Value = "Greatest % Decrease"
        wsheet.Range("O4").Value = "Greatest Total Volume"
        
        ' Format columns for width and percentage for better viewing
        wsheet.Columns("J").ColumnWidth = 12
        wsheet.Columns("K").ColumnWidth = 12
        wsheet.Columns("L").ColumnWidth = 15
        wsheet.Columns("O").ColumnWidth = 17
        wsheet.Range("J1:J" & TotalRowCount).NumberFormat = "0.00"
        wsheet.Range("K1:K" & TotalRowCount).NumberFormat = "0.00%"
        wsheet.Range("Q2:Q3").NumberFormat = "0.00%"
        wsheet.Range("Q4").NumberFormat = "0.00E+00"
        
        ' Initialize values for every worksheet for the second challenge
        GPI = 0
        GPD = 0
        TSV = 0
        
        ' Traverse through the rows in the current working sheet
        For Row = 2 To TotalRowCount
            ' First time TickerChanged is Forced to true so that we can initialize the values
            If TickerChanged Then
               OpenValue = wsheet.Cells(Row, 3).Value
               TotalStockVolume = 0
                PrintRow = PrintRow + 1
            End If
            
            TotalStockVolume = TotalStockVolume + wsheet.Cells(Row, 7).Value
            If (wsheet.Cells(Row, 1).Value <> wsheet.Cells(Row + 1, 1).Value) Then
                TickerChanged = True
                CloseValue = wsheet.Cells(Row, 6).Value
                YearlyChange = CloseValue - OpenValue
                PercentageChange = YearlyChange / OpenValue
                
                ' Print the values
                wsheet.Cells(PrintRow, 9).Value = wsheet.Cells(Row, 1).Value
                wsheet.Cells(PrintRow, 10).Value = YearlyChange
                If (YearlyChange < 0) Then
                    wsheet.Cells(PrintRow, 10).Interior.Color = vbRed
                Else
                    wsheet.Cells(PrintRow, 10).Interior.Color = vbGreen
                End If
                wsheet.Cells(PrintRow, 11).Value = PercentageChange
                wsheet.Cells(PrintRow, 12).Value = TotalStockVolume
                
                ' For the second challenge
                If PercentageChange > GPI Then
                   TickerGPI = wsheet.Cells(Row, 1).Value
                   GPI = PercentageChange
                   wsheet.Range("P2").Value = TickerGPI
                   wsheet.Range("Q2").Value = GPI
                End If
                
                If PercentageChange < GPD Then
                   TickerGPD = wsheet.Cells(Row, 1).Value
                   GPD = PercentageChange
                   wsheet.Range("P3").Value = TickerGPD
                   wsheet.Range("Q3").Value = GPD
                End If
                
                If TotalStockVolume > TSV Then
                   TickerTSV = wsheet.Cells(Row, 1).Value
                   TSV = TotalStockVolume
                   wsheet.Range("P4").Value = TickerTSV
                   wsheet.Range("Q4").Value = TSV
                End If
                
            Else
                TickerChanged = False
            End If
        
        Next Row
        ' End traversing across all rows in the current worksheet
        
    Next
    ' End searching across all worksheets
    
    EndTime = Timer
    MsgBox ("Completed in " & Format(EndTime - StartTime, "0.00") & " seconds")

End Sub
