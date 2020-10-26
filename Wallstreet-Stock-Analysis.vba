Sub WallStreet_Stock_Analysis()
    
    Dim aws As Worksheet

    'Start loop
    For Each aws In Worksheets

        'Create column labels for the summary table
        aws.Cells(1, 9).Value = "Ticker"
        aws.Cells(1, 10).Value = "Yearly Change"
        aws.Cells(1, 11).Value = "Percent Change"
        aws.Cells(1, 12).Value = "Total Stock Volume"

        'Set variable to hold the ticker symbol
        Dim ticker_symbol As String

        'Set variable to hold total volume of stock traded
        Dim total_vol As Double
        total_vol = 0

        Dim rowcount As Long
        rowcount = 2

        'Set variable to hold year open price
        Dim year_open As Double
        year_open = 0

        'Set variable to hold year close price
        Dim year_close As Double
        year_close = 0
        
        'Set variable to hold the change in price for the year
        Dim year_change As Double
        year_change = 0

        'Set variable to hold the percent change in price for the year
        Dim percent_change As Double
        percent_change = 0

        'Set variable for total rows to loop through
        Dim lastrow As Long
        lastrow = aws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop to search through ticker symbols
        For i = 2 To lastrow
            
            'Conditional to grab year open price
            If aws.Cells(i, 1).Value <> aws.Cells(i - 1, 1).Value Then
            year_open = aws.Cells(i, 3).Value

            End If

            'Total the volume for each row to determine the total stock volume for the year
            total_vol = total_vol + aws.Cells(i, 7)

            'Conditional to determine the ticker symbol is changing
            If aws.Cells(i, 1).Value <> aws.Cells(i + 1, 1).Value Then

                'Move ticker symbol to summary table
                aws.Cells(rowcount, 9).Value = aws.Cells(i, 1).Value

                'Move total stock volume to the summary table
                aws.Cells(rowcount, 12).Value = total_vol

                'Pull year end price
                year_close = aws.Cells(i, 6).Value

                'Calculate the price change for the year and move it to the summary table.
                year_change = year_close - year_open
                aws.Cells(rowcount, 10).Value = year_change

                'Conditional to format to highlight positive or negative change.
                If year_change >= 0 Then
                    aws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    aws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If

                'Calculate the percent change for the year and move it to the summary table format as a percentage
                'Conditional for calculating percent change
                If year_open = 0 And year_close = 0 Then
                    'Starting at zero and ending at zero will be a zero increase.  Cannot use a formula because
                    'it would be dividing by zero.
                    percent_change = 0
                    aws.Cells(rowcount, 11).Value = percent_change
                    aws.Cells(rowcount, 11).NumberFormat = "0.00%"
                ElseIf year_open = 0 Then
                    'If a stock starts at zero and increases, it grows by infinite percent.
                    '"New Stock" as percent change.
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    aws.Cells(rowcount, 11).Value = percent_change
                Else
                    percent_change = year_change / year_open
                    aws.Cells(rowcount, 11).Value = percent_change
                    aws.Cells(rowcount, 11).NumberFormat = "0.00%"
                End If

                'Add 1 to rowcount to move it to the next empty row in the summary table
                rowcount = rowcount + 1

                'Reset total stock volume, year open price, year close price, year change, year percent change
                total_vol = 0
                year_open = 0
                year_close = 0
                year_change = 0
                percent_change = 0
                
            End If
        Next i

        'Find best/worst performance table
        aws.Cells(2, 15).Value = "Greatest % Increase"
        aws.Cells(3, 15).Value = "Greatest % Decrease"
        aws.Cells(4, 15).Value = "Greatest Total Volume"
        aws.Cells(1, 16).Value = "Ticker"
        aws.Cells(1, 17).Value = "Value"

        'Find the lastrow of list of tickers
        lastrow = aws.Cells(Rows.Count, 9).End(xlUp).Row

        'Set variables to hold best performer, worst performer, and stock with the highest volume
        Dim best_stock As String
        Dim best_value As Double

        'Set best performer equal to the first stock
        best_value = aws.Cells(2, 11).Value

        Dim worst_stock As String
        Dim worst_value As Double

        'Set worst performer equal to the first stock
        worst_value = aws.Cells(2, 11).Value

        Dim most_vol_stock As String
        Dim most_vol_value As Double

        'Set most volume equal to the first stock
        most_vol_value = aws.Cells(2, 12).Value

        'Loop to search through summary table
        For j = 2 To lastrow

            'Conditional to determine best performer
            If aws.Cells(j, 11).Value > best_value Then
                best_value = aws.Cells(j, 11).Value
                best_stock = aws.Cells(j, 9).Value
            End If

            'Conditional to determine worst performer
            If aws.Cells(j, 11).Value < worst_value Then
                worst_value = aws.Cells(j, 11).Value
                worst_stock = aws.Cells(j, 9).Value
            End If

            'Conditional to determine stock with the greatest volume traded
            If aws.Cells(j, 12).Value > most_vol_value Then
                most_vol_value = aws.Cells(j, 12).Value
                most_vol_stock = aws.Cells(j, 9).Value
            End If

        Next j

        'Move best performer, worst performer, and stock with the highest volume items to the performance table
        aws.Cells(2, 16).Value = best_stock
        aws.Cells(2, 17).Value = best_value
        aws.Cells(2, 17).NumberFormat = "0.00%"
        aws.Cells(3, 16).Value = worst_stock
        aws.Cells(3, 17).Value = worst_value
        aws.Cells(3, 17).NumberFormat = "0.00%"
        aws.Cells(4, 16).Value = most_vol_stock
        aws.Cells(4, 17).Value = most_vol_value

        'Autofit table columns
        aws.Columns("I:L").EntireColumn.AutoFit
        aws.Columns("O:Q").EntireColumn.AutoFit

    Next aws

End Sub

