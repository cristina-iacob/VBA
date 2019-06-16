Sub StockInfo()

   Dim ws As Worksheet

    'Loop through all sheets
    For Each ws In Worksheets:
        ws.Activate

    'Create the variables for the stock data
    Dim TickerName As String
    Dim TotalVolume As Double
    Dim TableRow As Integer
    Dim LastRow As Long

   'Create headers for the new columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"

    'Set the initial value to the TotalVolume to 0
    TotalVolume = 0
   ' Create a variable to hold the current row value
    TableRow = 2

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

                   'Print the values for the new columns
                   Range("I" & TableRow).Value = TickerName
                   Range("J" & TableRow).Value = TotalVolume

                   'Go to the next line
                   TableRow = TableRow + 1
                   'Reset TotalVolume for the next ticker
                   TotalVolume = 0

                Else
                   'Add to the total volume value as you loop
                   TotalVolume = TotalVolume + Cells(i, 7).Value

               End If

            Next i

    Next ws
End Sub
