Sub NewVBAStock()
   
    'Set a variable for the worksheet
    Dim ws As Worksheet

    'Set the look for each worksheet
    For Each ws In Worksheets

        'Set the Ticker Name
        Dim Ticker_Name As String

        'Set a variable for the Price, Percent Change, Increase and Decrease in Percent
        Dim Price1 As Double
        Dim Price2 As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Increase_Name As String
        Dim Decrease_Name As String
        Dim Increase_Percent As Double
        Dim Decrease_Percent As Double
        Price1 = ws.Cells(2, 3).Value
        Price2 = 0
        Yearly_Change = 0
        Percent_Change = 0
        Increase_Percent = 0
        Decrease_Percent = 0

   
        'Set a Variable for Volume per Ticker and Total Volume
        Dim Ticker_Volume As Double
        Ticker_Volume = 0
        Dim Greatest_Volume_Name As String
        Dim Greatest_Volume As Double
        Greatest_Volume = 0

        'Set the variable for the last row
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Print Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 9).Font.Bold = True
        ws.Cells(1, 9).ColumnWidth = 10
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 10).Font.Bold = True
        ws.Cells(1, 10).ColumnWidth = 15
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 11).Font.Bold = True
        ws.Cells(1, 11).ColumnWidth = 15
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 12).Font.Bold = True
        ws.Cells(1, 12).ColumnWidth = 20
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 16).Font.Bold = True
        ws.Cells(1, 16).ColumnWidth = 10
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(1, 17).Font.Bold = True
        ws.Cells(1, 17).ColumnWidth = 15
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 15).ColumnWidth = 20
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 15).ColumnWidth = 20
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 15).ColumnWidth = 20

        'Track the rows
        Dim Ticker_Row As Long
        Ticker_Row = 2

        'Loop through the Tickers
        For i = 2 To lastrow
        
        'Check for different Ticker Name in the date field
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            Ticker_Name = ws.Cells(i, 1).Value
            'Print the Ticker Name
            ws.Range("I" & Ticker_Row).Value = Ticker_Name

            'Store the end price and Ticker Volume
            Price2 = Price2 + ws.Cells(i, 6).Value
            Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

            'Print the Ticker_Volume
            ws.Range("L" & Ticker_Row).Value = Ticker_Volume

            'Calculate the Ticker with the Greatest Volume
            If Ticker_Volume > Greatest_Volume Then
                Greatest_Volume = Ticker_Volume
                Greatest_Volume_Name = Ticker_Name
                
                'Print the Greatest Volume
                ws.Range("P4").Value = Greatest_Volume_Name
                ws.Range("Q4").Value = Greatest_Volume
            
            End If

            'Calculate the Price Change
            Yearly_Change = Price2 - Price1
                    
            'Print Yearly change in price
            ws.Range("J" & Ticker_Row).Value = Yearly_Change

            'Color the Yearly-Change
            If Yearly_Change < 0 Then
                ws.Range("J" & Ticker_Row).Interior.ColorIndex = 3
            Else
                ws.Range("J" & Ticker_Row).Interior.ColorIndex = 4
            End If

            'Calculate the Percent Price Change if open price is not 0
            If Price1 <> 0 Then
                Percent_Change = (Price2 - Price1) / Price1 * 100
                Percent_Change = Round(Percent_Change, 2)
   
                'Print percent change in price
                ws.Range("K" & Ticker_Row).Value = Percent_Change & "%"
                 
                'Calculate the greatest increase in percent change
                 If (Percent_Change > Increase_Percent) Then
                    Increase_Percent = Percent_Change
                    Increase_Percent = Round(Percent_Change, 2)
                    Increase_Name = Ticker_Name
                    
                End If

                'Calculate the greatest decrease in percent change
                If Percent_Change < Decrease_Percent Then
                    Decrease_Percent = Percent_Change
                    Decrease_Percent = Round(Percent_Change, 2)
                    Decrease_Name = Ticker_Name
                End If
                    
            Else
                'Calculate the Percent Price Change if open price is 0
                Percent_Change = Price2 / (Price1 - 1)
                Percent_Change = Round(Percent_Change, 2)
                
                'Print percent change in price
                ws.Range("K" & Ticker_Row).Value = Percent_Change & "%"
                
                'Calculate the greatest increase in percent change
                 If (Percent_Change > Increase_Percent) Then
                    Increase_Percent = Percent_Change
                    Increase_Percent = Round(Percent_Change, 2)
                    Increase_Name = Ticker_Name
                    
                End If

                'Calculate the greatest decrease in percent change
                If Percent_Change < Decrease_Percent Then
                    Decrease_Percent = Percent_Change
                    Decrease_Percent = Round(Percent_Change, 2)
                    Decrease_Name = Ticker_Name
                End If
                
            End If
                          
                'Print the Increase / Decrease in Percent Change
                ws.Range("P2").Value = Increase_Name
                ws.Range("Q2").Value = Increase_Percent & "%"
                ws.Range("P3").Value = Decrease_Name
                ws.Range("Q3").Value = Decrease_Percent & "%"

                
            'Add 1 to the Ticker Row
            Ticker_Row = Ticker_Row + 1

            'Reset the Price Fields
            Yearly_Change = 0
            Price1 = ws.Cells(i + 1, 3).Value
            Price2 = 0
            Ticker_Volume = 0

            'Check for 0101 in the date field
             Else: Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
       
            End If
        Next i
    Next ws
End Sub
'Found solution at https://freesoft.dev/program/163047389 used to verify my code