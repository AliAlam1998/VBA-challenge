Attribute VB_Name = "Module1"
Sub StockAnalysisAcrossWorksheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        StockAnalysis ws
    Next ws
End Sub

Sub StockAnalysis(ws As Worksheet)
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim Ticker As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim TotalVolume As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    
    Dim SummaryRow As Long
    SummaryRow = 2
    
    ' Insert column headers in the summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To LastRow
        ' Check if the ticker symbol has changed
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            YearOpen = ws.Cells(i, 3).Value
        End If
        
        ' Calculate total volume
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        ' Check if the next row has a different ticker symbol or is the last row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = LastRow Then
            Ticker = ws.Cells(i, 1).Value
            YearClose = ws.Cells(i, 6).Value
            
            ' Calculate yearly change and percent change
            YearlyChange = YearClose - YearOpen
            If YearOpen <> 0 Then
                PercentChange = (YearlyChange / YearOpen) * 100
            Else
                PercentChange = 0
            End If
            
            ' Output data to summary table
            ws.Cells(SummaryRow, 9).Value = Ticker
            ws.Cells(SummaryRow, 10).Value = YearlyChange
            ws.Cells(SummaryRow, 11).Value = PercentChange
            ws.Cells(SummaryRow, 12).Value = TotalVolume
            
            ' Format percent change column
            If YearlyChange > 0 Then
                ws.Cells(SummaryRow, 11).Interior.Color = RGB(0, 255, 0)
            ElseIf YearlyChange < 0 Then
                ws.Cells(SummaryRow, 11).Interior.Color = RGB(255, 0, 0)
            Else
                ws.Cells(SummaryRow, 11).Interior.Color = RGB(255, 255, 255)
            End If
            
            ' Find greatest % increase, % decrease, and total volume
            If PercentChange > GreatestIncrease Then
                GreatestIncrease = PercentChange
                GreatestIncreaseTicker = Ticker
            ElseIf PercentChange < GreatestDecrease Then
                GreatestDecrease = PercentChange
                GreatestDecreaseTicker = Ticker
            End If
            If TotalVolume > GreatestVolume Then
                GreatestVolume = TotalVolume
                GreatestVolumeTicker = Ticker
            End If
            
            ' Reset variables for the next ticker
            YearOpen = 0
            YearlyChange = 0
            PercentChange = 0
            TotalVolume = 0
            
            ' Move to the next summary row
            SummaryRow = SummaryRow + 1
        End If
    Next i
    
    ' Autofit columns
    ws.Columns("I:Q").AutoFit
    
    ' Output greatest % increase, % decrease, and total volume
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 16).Value = GreatestIncreaseTicker
    ws.Cells(3, 16).Value = GreatestDecreaseTicker
    ws.Cells(4, 16).Value = GreatestVolumeTicker
    ws.Cells(2, 17).Value = FormatPercent(GreatestIncrease / 100, 2)
    ws.Cells(3, 17).Value = FormatPercent(GreatestDecrease / 100, 2)
    ws.Cells(4, 17).Value = GreatestVolume
    
    ' Format greatest % increase and % decrease values
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ' Format total volume as a number
    ws.Cells(4, 17).NumberFormat = "#,##0"
End Sub

