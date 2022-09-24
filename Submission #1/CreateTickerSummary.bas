Attribute VB_Name = "Module1"
Sub CreateTickerSummary()
'Loop through all stocks for one year and output each ticker symbol with yearly and percent change and total stock volume.

    Dim i As Long, j As Long
    Dim wsCount As Long: wsCount = ThisWorkbook.Worksheets.Count

    For i = 1 To wsCount    'Loop through worksheets.
        Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(i)

        With ws
            'Set up ticker summary table.
            .Cells(1, 9).Value2 = "Ticker"
            .Cells(1, 10).Value2 = "Yearly Change"
            .Cells(1, 11).Value2 = "Percent Change"
            .Cells(1, 12).Value2 = "Total Stock Volume"

            Dim LastSummaryRow As Long: LastSummaryRow = 2

            'Store the very first open price and volume value.
            Dim OpenPrice As Double: OpenPrice = .Cells(2, 3).Value2
            Dim Volume As Double: Volume = .Cells(2, 7).Value2  'Used double instead of long/integer because of overflow run-time error.

            Dim LastRow As Long: LastRow = .Range("A1").End(xlDown).Row
            For j = 2 To LastRow    'Loop through tickers.
                If .Cells(j, 1).Value2 = .Cells(j + 1, 1).Value2 Then 'Check if next ticker from current row is the same.
                    Volume = Volume + .Cells(j + 1, 7).Value2
                Else    'If the next ticker from the current row is different, output the ticker's summary information.
                    .Cells(LastSummaryRow, 9).Value2 = .Cells(j, 1).Value2  'Output ticker symbol.

                    Dim rngYearlyChange As Range: Set rngYearlyChange = .Cells(LastSummaryRow, 10)
                    rngYearlyChange.Value2 = .Cells(j, 6).Value2 - OpenPrice 'Output yearly change (i.e., latest close price minus earliest open price).
                    If rngYearlyChange.Value2 < 0 Then
                        rngYearlyChange.Interior.ColorIndex = 3    'Change fill color to red if yearly change is negative.
                    Else
                        rngYearlyChange.Interior.ColorIndex = 4    'Change fill color to green if yearly change is positive.
                    End If

                    Dim rngPercentChange As Range: Set rngPercentChange = .Cells(LastSummaryRow, 11)
                    rngPercentChange.Value2 = rngYearlyChange.Value2 / OpenPrice 'Output percent change (i.e., yearly change divided by open price).
                    rngPercentChange.NumberFormat = "0.00%" 'Format Percent Change to percents.

                    .Cells(LastSummaryRow, 12).Value2 = Volume 'Output volume.

                    'Prep for next ticker.
                    If j <> LastRow Then
                        OpenPrice = .Cells(j + 1, 3).Value2 'Retrieve open price of next ticker.
                        Volume = .Cells(j + 1, 7) 'Reset volume.
                        LastSummaryRow = LastSummaryRow + 1
                    End If
                End If
            Next j

            .Columns("I:L").AutoFit 'Autofit summary table columns.

'BONUS: Output stocks with Greatest % Increase; Greatest % Decrease; and Greatest Total Volume.
            'Set up bonus table.
            .Cells(1, 14).Value2 = .Name
            .Cells(1, 15).Value2 = "Ticker"
            .Cells(1, 16).Value2 = "Value"
            .Cells(2, 14).Value2 = "Greatest % Increase"
            .Cells(3, 14).Value2 = "Greatest % Decrease"
            .Cells(4, 14).Value2 = "Greatest Volume"

            LastRow = .Range("I1").End(xlDown).Row
            Dim MaxPCTicker As String, MaxPC As Double, MinPCTicker As String, MinPC As Double, MaxVolTicker As String, MaxVol As Double
            For j = 2 To LastRow    'Loop through newly created ticker summary table.
                'Check if Percent Change is either greater or lesser than last.
                If .Cells(j, 11).Value2 > MaxPC Then
                    MaxPCTicker = .Cells(j, 9).Value2
                    MaxPC = .Cells(j, 11).Value2
                ElseIf .Cells(j, 11).Value2 < MinPC Then
                    MinPCTicker = .Cells(j, 9).Value2
                    MinPC = .Cells(j, 11).Value2
                End If

                'Check if Volume is greater than last.
                If .Cells(j, 12).Value2 > MaxVol Then
                    MaxVolTicker = .Cells(j, 9).Value2
                    MaxVol = .Cells(j, 12).Value2
                End If
            Next j

            'Store results into bonus table.
            .Cells(2, 15).Value2 = MaxPCTicker
            .Cells(2, 16).Value2 = MaxPC
            .Cells(3, 15).Value2 = MinPCTicker
            .Cells(3, 16).Value2 = MinPC
            .Cells(4, 15).Value2 = MaxVolTicker
            .Cells(4, 16).Value2 = MaxVol

            .Range("P2:P3").NumberFormat = "0.00%"  'Format Greatest % Increase and Greatest % Decrease to percents.

            .Columns("N:P").AutoFit 'Autofit bonus table columns.
        End With
    Next i
End Sub
