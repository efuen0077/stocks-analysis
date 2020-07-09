Attribute VB_Name = "Module10"

Sub AllStocksAnalysis()

    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("VBA of Wallstreet Challenge").Activate

    
    Range("A1").Value = "All Stocks (" + yearValue + ")"


   'Create a header row

   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"
   
   'We need to declare the tickers array and 3 other arrays.

    Dim tickers(12) As String
    Dim volume(12) As String

   'we need to create a tickerIndex variable in order to even set it to zero.

    Dim tickerIndex As Integer


    Dim startingPrices(12) As String
    Dim endingPrices(12) As String
   
   Worksheets(yearValue).Activate

   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'let's set the tickerIndex to zero.
   tickerIndex = 0

   Worksheets(yearValue).Activate
   
   
'We are going to have a series of nested loops within our main for loop
   For tickerIndex = 0 To 11

       Worksheets(yearValue).Activate

       For j = 2 To RowCount

           'the following "if" statement retrieves the names and starting values and stores them in
           'the arrays that we created earlier.
        
           If Cells(j, 1).Value <> Cells(j - 1, 1).Value Then

               tickers(tickerIndex) = Cells(j, 1).Value

               startingPrices(tickerIndex) = Cells(j, 6).Value
               
           End If

               Worksheets(yearValue).Activate

                   TotalVolume = 0

                   For e = 2 To RowCount

                       If Cells(e, 1).Value = tickers(tickerIndex) Then

                           TotalVolume = TotalVolume + Cells(e, 8).Value

                       End If

                   Next e

                   volume(tickerIndex) = TotalVolume

           'retrieve and store ending price in array as well as increment tickerIndex for next loop


           If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then

               endingPrices(tickerIndex) = Cells(j, 6).Value

               tickerIndex = tickerIndex + 1

           End If

       Next j

   Next tickerIndex

   'We want all of this info to go to our worksheet "VBA of Wallstreet Challenge"

   Worksheets("VBA of Wallstreet Challenge").Activate


   For i = 0 To 11

       Cells(i + 4, 1).Value = tickers(i)

       Cells(i + 4, 3).Value = endingPrices(i) / startingPrices(i) - 1

       Cells(4 + i, 2).Value = volume(i)

   Next i


Worksheets("VBA of Wallstreet Challenge").Activate

       Range("A3:C3").Font.Bold = True
       Range("A1").Font.FontStyle = "Bold"
       Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
       Range("B4:B15").NumberFormat = "#,##0"
       Range("c4:c15").NumberFormat = "0.0%"
       Columns(2).AutoFit


   'color our cells


   Worksheets("VBA of Wallstreet Challenge").Activate

   dataRowStart = 4
   dataRowEnd = Cells(Rows.Count, "C").End(xlUp).Row

   For g = dataRowStart To dataRowEnd


       If Cells(g, 3).Value > 0 Then


           Cells(g, 3).Interior.Color = vbGreen


       ElseIf Cells(g, 3).Value < 0 Then

'Let's make sure that the cell is magenta (module used red, but I'd like to customize a little more.

           Cells(g, 3).Interior.Color = vbMagenta
       Else
           Cells(g, 3).Interior.Color = xlNone

       End If

   Next g

End Sub

