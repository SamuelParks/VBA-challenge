Attribute VB_Name = "Module1"
'Runs through each worksheet doing each analysis
Sub StockDataAnalysisAllWorksheets()

Dim ws As Worksheet

'A website said that the line below makes VBA run faster through multiple worksheets
'Application.ScreenUpdating = False
    
    For Each ws In Worksheets
        ws.Select
        Call StockDataAnalysis
    Next ws

'A website said that the line below makes VBA run faster through multiple worksheets
'Application.ScreenUpdating = True
 
    MsgBox ("Completed analysis for each worksheet.")
    
End Sub

'This does the analysis and the output on a given worksheet
Sub StockDataAnalysis()

    'Declaring variables
  
    Dim SummaryTableRow As Integer
    
    Dim Col1 As String
    Dim Col2 As Double
    Dim Col3 As Double
    Dim Col4 As Double
    Dim Col5 As Double
    Dim Col6 As Double
    Dim Col7 As Double
    
    Dim TickerSymbol As String
    Dim YearBeginningStockDate As Double
    Dim YearBeginningLowPrice As Double
    Dim YearBeginningOpenPrice As Double
    Dim YearBeginningHighPrice As Double
    Dim YearBeginningClosePrice As Double
    Dim Vol As Double
    
    Dim YearEndStockDate As Double
    Dim YearEndLowPrice As Double
    Dim YearEndOpenPrice As Double
    Dim YearEndHighPrice As Double
    Dim YearEndClosePrice As Double
    
    Dim YearlyChange As Double
    Dim YearlyPercentChange As Double


    SummaryTableRow = 2
         
         
        TickerSymbol = "Empty"
        YearBeginningStockDate = Cells(2, 2).Value
        ' YearBeginningLowPrice = 0
        YearBeginningOpenPrice = Cells(2, 3).Value
        '  YearBeginningHighPrice = 0
        ' YearBeginningClosePrice = 0
        
        YearEndStockDate = 0
        ' YearEndLowPrice = 0
        ' YearEndOpenPrice = 0
        ' YearEndHighPrice = 0
        YearEndClosePrice = 0
        Vol = 0
        
        
    'I got this from the StarCounter Bonus: This counts the number of rows
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  
    ' Loop through all rows on tab except the header
    For i = 2 To lastRow

        ' What is in column 1
        Col1 = Cells(i, 1).Value
        TickerSymbol = Col1
        
        ' What is in column 2
        Col2 = Cells(i, 2).Value
                
        ' What is in column 3
        Col3 = Cells(i, 3).Value
         
         
     ' Col 4 and 5 information is not needed for this assignment
        ' What is in column 4
       ' Col4 = Cells(i, 4).Value
        
        ' What is in column 5
      '  Col5 = Cells(i, 5).Value
        
        ' What is in column 6
        Col6 = Cells(i, 6).Value
        
        ' What is in column 7
        Col7 = Cells(i, 7).Value
        
        'Sets beginning of year information for Stock
        If Col2 < EarliestStockDate Then
            YearBeginningStockDate = Col2
            YearBeginningOpenPrice = Col3
         '  YearBeginningHighPrice = Col4
         '  YearBeginningLowPrice = Col5
         '  YearBeginningClosePrice = Col6
                         
        'Sets end of year information for Stock
        ElseIf Col2 > LatestStockDate Then
            YearEndDate = Col2
           ' YearEndOpenPrice = Col3
           ' YearEndHighPrice = Col4
           ' YearEndLowPrice = Col5
            YearEndClosePrice = Col6
        End If

   'Adds up volume for each row cumulatively
   Vol = Vol + Col7
   
   
    ' Searches for when the value of the next TickerSymbol cell is different than that of the current TickerSymbol cell for the row
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
     
     'Print Ticker Information Output in the Summary Table
      Range("I" & SummaryTableRow).Value = TickerSymbol

      ' Print the Yearly Change to the Summary Table
      YearlyChange = YearEndClosePrice - YearBeginningOpenPrice
      Range("J" & SummaryTableRow).Value = YearlyChange
      
            'Checks to see if there was no the beginning price was zero to avoid Divide by Zero issue
            If (YearBeginningOpenPrice <> 0) Then
        
            ' Print the Yearly Percent Change to the Summary Table
             YearlyPercentChange = YearlyChange / YearBeginningOpenPrice
             Range("K" & SummaryTableRow).Value = YearlyPercentChange
             
            Else
                YearlyPercentChange = 0
                Range("K" & SummaryTableRow).Value = "Divide by Zero - Exclude due to beginning price equalling zero."
             End If

      ' Print the Total Volume to the Summary Table
            Range("L" & SummaryTableRow).Value = Vol

      ' Add one to the summary table row
      SummaryTableRow = SummaryTableRow + 1
      
      
      
    ' Reset Variable for next symbol

        TickerSymbol = "Empty"
        YearBeginningStockDate = Cells(i + 1, 2).Value
        ' YearBeginningLowPrice = 0
        YearBeginningOpenPrice = Cells(i + 1, 3).Value
        ' YearBeginningHighPrice = 0
        ' YearBeginningClosePrice = 0
        
        YearEndStockDate = 0
        ' YearEndLowPrice = 0
        ' YearEndOpenPrice = 0
        ' YearEndHighPrice = 0
        YearEndClosePrice = 0
        Vol = 0

    End If


    
    ' Call the next iteration
    Next i

Range("I" & "1").Value = "Stock Ticker"
Range("J" & "1").Value = "Yearly Change"
Range("K" & "1").Value = "Percentage Yearly Change"
Range("L" & "1").Value = "Total Stock Volume Traded"

'This Formats the Yearly Change Column - 10 is entered for column J = 10th column
FormatYearlyChange (10)

'This produces the "greatest" grid - starts are PercentageYearlyChange Column - 11 is entered for column K = 11th column
TheGreatestForYear (11)

End Sub

'This formats the Yearly Changes Summary Area
Sub FormatYearlyChange(columnNumber)

'I got this from the StarCounter Bonus: This counts the number of rows
    lastRow = Cells(Rows.Count, columnNumber).End(xlUp).Row

    ' Loop through all rows on tab except the header
    For i = 2 To lastRow
    
    Cells(i, columnNumber).NumberFormat = "0.00"
    Cells(i, columnNumber + 1).NumberFormat = "0.00%"
    Cells(i, columnNumber + 2).NumberFormat = "0,000"
    
        If (Cells(i, columnNumber).Value < 0) Then
             
             ' Set the Cell Colors to Red
      Cells(i, columnNumber).Interior.ColorIndex = 3
    
        ElseIf (Cells(i, columnNumber).Value > 0) Then
      ' Set the Cell Colors to Green
      Cells(i, columnNumber).Interior.ColorIndex = 4
      
        End If
        
    Next i
    
End Sub

'This produces the "greatest" grid - starts are PercentageYearlyChange Column
Sub TheGreatestForYear(PercYearlyChangeColumnNumber)

Dim GreatestIncreaseVal As Double
Dim GreatestDecreaseVal As Double
Dim GreatestVolVal As Double

Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestVolTicker As String

'I got this from the StarCounter Bonus: This counts the number of rows
    lastRow = Cells(Rows.Count, PercYearlyChangeColumnNumber).End(xlUp).Row

    For i = 2 To lastRow
    
        'Checks to make sure a cell has a number and skips if it is not
        If IsNumeric(Cells(i, PercYearlyChangeColumnNumber).Value) = True Then
        
            'This tracks the greatest % increase and the accompanying ticker as it goes through the rows
            If Cells(i, PercYearlyChangeColumnNumber).Value > GreatestIncreaseVal Then
                GreatestIncreaseVal = Cells(i, PercYearlyChangeColumnNumber).Value
                GreatestIncreaseTicker = Cells(i, PercYearlyChangeColumnNumber - 2).Value
            End If
            
            'This tracks the greatest % decrease and the accompanying ticker as it goes through the rows
            If Cells(i, PercYearlyChangeColumnNumber).Value < GreatestDecreaseVal Then
                GreatestDecreaseVal = Cells(i, PercYearlyChangeColumnNumber).Value
                GreatestDecreaseTicker = Cells(i, PercYearlyChangeColumnNumber - 2).Value
            End If
            
         End If
            
         'Checks to make sure a cell has a number and skips if it is not
        If IsNumeric(Cells(i, PercYearlyChangeColumnNumber + 1).Value) = True Then
            
            'This tracks the greatest volume and the accompanying ticker as it goes through the rows
            If Cells(i, PercYearlyChangeColumnNumber + 1).Value > GreatestVolVal Then
                GreatestVolVal = Cells(i, PercYearlyChangeColumnNumber + 1).Value
                GreatestVolTicker = Cells(i, PercYearlyChangeColumnNumber - 2).Value
            End If
            
        End If
    
    Next i
    
'This outputs a summary of the greatest
    Cells(1, PercYearlyChangeColumnNumber + 5).Value = "Ticker"
    Cells(1, PercYearlyChangeColumnNumber + 6).Value = "Value"

    Cells(2, PercYearlyChangeColumnNumber + 4).Value = "Greatest % Increase"
    Cells(2, PercYearlyChangeColumnNumber + 5).Value = GreatestIncreaseTicker
    Cells(2, PercYearlyChangeColumnNumber + 6).Value = GreatestIncreaseVal
    Cells(2, PercYearlyChangeColumnNumber + 6).NumberFormat = "0.00%"

    
    Cells(3, PercYearlyChangeColumnNumber + 4).Value = "Greatest % Decrease"
    Cells(3, PercYearlyChangeColumnNumber + 5).Value = GreatestDecreaseTicker
    Cells(3, PercYearlyChangeColumnNumber + 6).Value = GreatestDecreaseVal
    Cells(3, PercYearlyChangeColumnNumber + 6).NumberFormat = "0.00%"
    
    Cells(4, PercYearlyChangeColumnNumber + 4).Value = "Greatest Total Volume"
    Cells(4, PercYearlyChangeColumnNumber + 5).Value = GreatestVolTicker
    Cells(4, PercYearlyChangeColumnNumber + 6).Value = GreatestVolVal
    Cells(4, PercYearlyChangeColumnNumber + 6).NumberFormat = "0,000"
    
End Sub



