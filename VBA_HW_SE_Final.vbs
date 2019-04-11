Sub Stocks()

' Set an initial variable for the worksheet
Dim ws As Worksheet

' Set a variable for the starting worksheet
Dim starting_ws As Worksheet

' Set the starting sheet as the sheet that's active
Set starting_ws = ActiveSheet

    ' Loop through each year's sheete
    For Each ws In ThisWorkbook.Worksheets

       ' Activate the worksheet 
        ws.Activate

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set the header for the ticker column of the summary
        ws.Cells(1, 9).Value = "Ticker"

        ' Set the header for the total column of the summary 
        ws.Cells(1, 10).Value = "Total Stock Volume"

        ' Set an initial variable for holding the ticker name
        Dim Ticker As String

        ' Set an initial variable for holding the total stock volume per ticker
        Dim Ticker_Total As Double
        Ticker_Total = 0

        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Loop through all ticker entries
        For i = 2 To LastRow

            ' Check if we are still within the ticker, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                ' Set the ticker name
                Ticker = Cells(i, 1).Value

                ' Add to the ticker total
                Ticker_Total = Ticker_Total + Cells(i, 7).Value

                ' Print the ticker in the Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker

                ' Print the total stock volume to the Summary Table
                Range("J" & Summary_Table_Row).Value = Ticker_Total

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
        
                ' Reset the ticker total
                Ticker_Total = 0

            ' If the cell immediately following a row is the same ticker...
            Else

            ' Add to the ticker total
            Ticker_Total = Ticker_Total + Cells(i, 7).Value

            End If

        Next i

    Next
    
    ' Return to starting sheet
    starting_ws.Activate

End Sub