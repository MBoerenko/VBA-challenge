Sub BasicVBAStock()
   
        'Set the Ticker Name
        Dim Ticker_Name As String

        'Set a variable for the Price, Percent Change, Increase and Decrease in Percent
        Dim Price1 As Double
        Dim Price2 As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Price1 = Cells(2, 3).Value
        Price2 = 0
        Yearly_Change = 0
        Percent_Change = 0
   
        'Set a Variable for Volume per Ticker and Total Volume
        Dim Ticker_Volume As Double
        Ticker_Volume = 0

        'Set the variable for the last row
        Dim lastrow As Long
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Print Headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 9).Font.Bold = True
        Cells(1, 9).ColumnWidth = 10
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 10).Font.Bold = True
        Cells(1, 10).ColumnWidth = 15
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 11).Font.Bold = True
        Cells(1, 11).ColumnWidth = 15
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 12).Font.Bold = True
        Cells(1, 12).ColumnWidth = 20

        'Track the rows
        Dim Ticker_Row As Long
        Ticker_Row = 2

        'Loop through the Tickers
        For i = 2 To lastrow
        
        'Check for different Ticker Name in the date field
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            Ticker_Name = Cells(i, 1).Value
            'Print the Ticker Name
            Range("I" & Ticker_Row).Value = Ticker_Name

            'Store the end price and Ticker Volume
            Price2 = Price2 + Cells(i, 6).Value
            Ticker_Volume = Ticker_Volume + Cells(i, 7).Value

            'Print the Ticker_Volume
            Range("L" & Ticker_Row).Value = Ticker_Volume

            'Calculate the Price Change
            Yearly_Change = Price2 - Price1
                    
            'Print Yearly change in price
            Range("J" & Ticker_Row).Value = Yearly_Change

            'Color the Yearly-Change
            If Yearly_Change < 0 Then
                Range("J" & Ticker_Row).Interior.ColorIndex = 3
            Else
                Range("J" & Ticker_Row).Interior.ColorIndex = 4
            End If

            'Calculate the Percent Price Change if open price is not 0
            If Price1 <> 0 Then
                Percent_Change = (Price2 - Price1) / Price1 * 100
                Percent_Change = Round(Percent_Change, 2)
   
                'Print percent change in price
                Range("K" & Ticker_Row).Value = Percent_Change & "%"
                 
            Else
                'Calculate the Percent Price Change if open price is 0
                Percent_Change = Price2 / (Price1 - 1)
                Percent_Change = Round(Percent_Change, 2)
                
                'Print percent change in price
                Range("K" & Ticker_Row).Value = Percent_Change & "%"
                
            End If
          
            'Add 1 to the Ticker Row
            Ticker_Row = Ticker_Row + 1

            'Reset the Price Fields
            Yearly_Change = 0
            Price1 = Cells(i + 1, 3).Value
            Price2 = 0
            Ticker_Volume = 0

            'Check for 0101 in the date field
             Else: Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
       
            End If
        Next i
End Sub
'Found solution at https://freesoft.dev/program/163047389 used to verify my code