Sub StockInfo()

   Dim ws As Worksheet

    'Loop through all sheets
    For Each ws In Worksheets:
        ws.Activate

    'Create the variables for the stock data
    Dim TickerName As String
    Dim TotalVolume As Double
    Dim CurrentRow As Integer
    Dim LastRow As Long
    Dim OpeningValue As Double
    Dim ClosingValue As Double

    'Set the first value for Yearly Change as the first Opening Value
     OpeningValue = Range("C2").Value

   'Create headers for the new columns and format the percent change column
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    'Range("J2:J" & LastRow).NumberFormat = "0.000000000"
    Range("K1").Value = "Percent Change"
    'Range("K2:K" & LastRow).NumberFormat = "0.00%"
    Range("L1").Value = "Total Stock Volume"

    'Set the initial value to the TotalVolume to 0
    TotalVolume = 0
   ' Create a variable to hold the current row value
    CurrentRow = 2

    'Determine the Last Row
    LastRow = Range("A1").End(xlDown).Row

               'Loop through all rows
               For i = 2 To LastRow

               'Check if it's still within the same ticker
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                   'Set the ticker
                   TickerName = Cells(i, 1).Value
                   'Increasing TotalVolume by current rows volume
                   TotalVolume = TotalVolume + Cells(i, 7).Value
                   'set ClosingValue
                   ClosingValue = Cells(i, 6)
                   'Calculate YearlyChange
                   YearlyChange = ClosingValue - OpeningValue

                       'Eliminate the situation of division by zero when Opening Value is 0
                        If OpeningValue = 0 Then
                            PercentChange = 0
                        Else
                            PercentChange = (YearlyChange / OpeningValue) * 100
                        End If

                       'Print the values for the new columns
                        Range("I" & CurrentRow).Value = TickerName
                        Range("J" & CurrentRow).Value = YearlyChange
                        'Range("J" & CurrentRow).NumberFormat = "0.000000000000000”
                        Range("K" & CurrentRow).Value = PercentChange & "%"
                        Range("L" & CurrentRow).Value = TotalVolume

                   'Set next opening value
                   OpeningValue = Cells(i + 1, 3)
                   'Reset TotalVolume for the next ticker
                   TotalVolume = 0

                        'Positive yearly change --> fill column with Green
                        If Range("J" & CurrentRow).Value >= 0 Then
                                   Range("J" & CurrentRow).Interior.ColorIndex = 4
                        'Negative yearly change --> Red
                        ElseIf Range("J" & CurrentRow).Value < 0 Then
                                   Range("J" & CurrentRow).Interior.ColorIndex = 3
                        End If


                   'Go to the next line
                   CurrentRow = CurrentRow + 1

                Else
                   'Add to the total volume value as you loop
                   TotalVolume = TotalVolume + Cells(i, 7).Value

               End If

            Next i

    Next ws
End Sub
