VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub QuarterlyStockAnalysis()

    Dim ws As Worksheet
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim TotalVolume As Double
    Dim QuarterlyChange As Double
    Dim PercentageChange As Double
    Dim LastRow As Long
    Dim SummaryRow As Long
    Dim i As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim TickerGreatestIncrease As String
    Dim TickerGreatestDecrease As String
    Dim TickerGreatestVolume As String

    For Each ws In Worksheets
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        SummaryRow = 2
        
        With ws
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Quarterly Change"
            .Cells(1, 11).Value = "Percentage Change"
            .Cells(1, 12).Value = "Total Volume"
        End With
        
        Ticker = ws.Cells(2, 1).Value
        OpenPrice = ws.Cells(2, 3).Value
        TotalVolume = 0
        
        GreatestIncrease = -1E+30
        GreatestDecrease = 1E+30
        GreatestVolume = 0
        
        For i = 2 To LastRow
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value

            If ws.Cells(i + 1, 1).Value <> Ticker Or i = LastRow Then
                ClosePrice = ws.Cells(i, 6).Value
                QuarterlyChange = ClosePrice - OpenPrice
                PercentageChange = IIf(OpenPrice <> 0, QuarterlyChange / OpenPrice, 0)

                With ws
                    .Cells(SummaryRow, 9).Value = Ticker
                    .Cells(SummaryRow, 10).Value = QuarterlyChange
                    .Cells(SummaryRow, 11).Value = PercentageChange
                    .Cells(SummaryRow, 11).NumberFormat = "0.00%"
                    .Cells(SummaryRow, 12).Value = TotalVolume
                    .Cells(SummaryRow, 12).NumberFormat = "0"
                    
                    If QuarterlyChange > 0 Then
                        .Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                    ElseIf QuarterlyChange < 0 Then
                        .Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                    End If
                End With
                
                If PercentageChange > GreatestIncrease Then
                    GreatestIncrease = PercentageChange
                    TickerGreatestIncrease = Ticker
                End If
                
                If PercentageChange < GreatestDecrease Then
                    GreatestDecrease = PercentageChange
                    TickerGreatestDecrease = Ticker
                End If
                
                If TotalVolume > GreatestVolume Then
                    GreatestVolume = TotalVolume
                    TickerGreatestVolume = Ticker
                End If
                
                SummaryRow = SummaryRow + 1
                
                If i <> LastRow Then
                    Ticker = ws.Cells(i + 1, 1).Value
                    OpenPrice = ws.Cells(i + 1, 3).Value
                    TotalVolume = 0
                End If
            End If
        Next i

        With ws
            .Cells(1, 14).Value = ""
            .Cells(1, 15).Value = "Ticker"
            .Cells(1, 16).Value = "Values"

            .Cells(2, 14).Value = "Greatest % Increase"
            .Cells(2, 15).Value = TickerGreatestIncrease
            .Cells(2, 16).Value = GreatestIncrease
            
            .Cells(3, 14).Value = "Greatest % Decrease"
            .Cells(3, 15).Value = TickerGreatestDecrease
            .Cells(3, 16).Value = GreatestDecrease
            
            .Cells(4, 14).Value = "Greatest Total Volume"
            .Cells(4, 15).Value = TickerGreatestVolume
            .Cells(4, 16).Value = GreatestVolume

            .Cells(2, 16).NumberFormat = "0.00%"
            .Cells(3, 16).NumberFormat = "0.00%"
            .Cells(4, 16).NumberFormat = "0.00E+00"
        End With
        
    Next ws

End Sub
