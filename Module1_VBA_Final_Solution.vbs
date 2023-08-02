Attribute VB_Name = "Module1_VBA_Final_Solution"
Option Explicit

' This is to tabulate a given dataset of stock prices from Year 2018-2020 in separate Worksheet
' Compiled the dataset from ColumnI to L for each TickerName followed by
' Yearly Change (lastClosingPrice - 1stOpenPrice) and the Percent Change
' Finally Total Stock Voume.
' Conditional formating applied for Yearly Change and Percent Change (Green for > 0 and Red for < 0)
'
' Another tabulation done in Column O to Q for Greatest % increase. Greatest % Decrease
' and Greatest Colume with corresponding Ticker and Numbers
' Autofit all the corresponding Columns


Sub Module1_VBA_Final_Solution()

    ' Initializing all variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRowColumnA As Long
    Dim lastRowColumnI As Long
    Dim outputRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim i As Long
    
    Dim maxPercentageIncrease As Double
    Dim maxPercentageDecrease As Double
    Dim maxTotalVolume As Double
    Dim tickerMaxPercentageIncrease As String
    Dim tickerMaxPercentageDecrease As String
    Dim tickerMaxTotalVolume As String

    
    
    Set wb = ThisWorkbook
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In wb.Worksheets
    
        ' Add header to the output columns
         ws.Cells(1, 9).Value = "Ticker"
         ws.Cells(1, 10).Value = "Yearly Change"
         ws.Cells(1, 11).Value = "Percent Change"
         ws.Cells(1, 12).Value = "Total Stock Volume"
         
   
         ' Find the last row in column A
         lastRowColumnA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ' lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
             
         ' Initialize outputRow to 2 (assuming header is in row 1)
         outputRow = 2
         
         totalVolume = 0 'set initial totalVolume to 0
             
         ' Loop through the data based on Column A
         For i = 2 To lastRowColumnA ' Assuming the data starts from row 2 and has headers in row 1
                 
             ' Get the current tickerName
             ticker = ws.Cells(i, 1).Value
                 
             ' Check if it's a new TickerName
             If ws.Cells(i - 1, 1).Value <> ticker Then
                 ' If it's a new TickerName, extract OpeningPrice
                 openingPrice = ws.Cells(i, 3).Value
             End If
                 
             ' Update ClosingPrice for each row (it will keep getting updated until the last occurrence of the TickerName)
             closingPrice = ws.Cells(i, 6).Value
                 
             ' Check if it's the last occurrence of the TickerName
             If ws.Cells(i + 1, 1).Value <> ticker Then
                                  
                 ' Add Volume to the Total Volume
                 totalVolume = totalVolume + ws.Cells(i, 7).Value
                     
                 ' If it's the last occurrence, output the data to columns I to L
                 ' line by line controlled by outputRow
                 ws.Cells(outputRow, 9).Value = ticker ' TickerName
                 ws.Cells(outputRow, 10).Value = closingPrice - openingPrice ' YearlyChange
                 ws.Cells(outputRow, 11).Value = (closingPrice - openingPrice) / openingPrice ' PercentChange
                 ws.Cells(outputRow, 12).Value = totalVolume ' TotalVolume
                 
                 ' Format PercentChange output to percentage
                 ws.Cells(outputRow, 11).NumberFormat = "0.00%" ' Format to percentage
                 
                     
                 ' Increment outputRow for the next TickerName
                 outputRow = outputRow + 1
                 
                 totalVolume = 0 'reset totalVolume to 0 for next TickerName
                 
             Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value ' TotalVolume
                 
             End If
             
         Next i
         
        ' -----------------------------------------------------------------------------
        ' CONDITIONAL FORMATTING (Green for positives value and Red for negative values
        ' -----------------------------------------------------------------------------

        ' Find the last row with data in column I
        lastRowColumnI = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        ' Loop through each row in column I
        For i = 2 To lastRowColumnI ' Assuming row 1 is the header
            ' Check if the value in column J is negative or positive
            If ws.Cells(i, "J").Value < 0 Then
                ' Format the cell with red color for negative values
                ws.Cells(i, "J").Interior.Color = RGB(255, 0, 0)
            Else
                ' Format the cell with green color for positive values
                ws.Cells(i, "J").Interior.Color = RGB(0, 255, 0)
            End If
                
            ' Check if the value in column K is negative or positive
            If ws.Cells(i, "K").Value < 0 Then
                ' Format the cell with red color for negative values
                ws.Cells(i, "K").Interior.Color = RGB(255, 0, 0)
            Else
                ' Format the cell with green color for positive values
                ws.Cells(i, "K").Interior.Color = RGB(0, 255, 0)
            End If
            
        Next i
        
        
        ' --------------------------------------------
        ' FIND GREATEST PERCENTAGE CHANGE AND VOLUME
        ' --------------------------------------------
    
        ' Initialize the variables to store the maximum and minimum values and corresponding tickers
        maxPercentageIncrease = ws.Cells(2, "K").Value
        maxPercentageDecrease = ws.Cells(2, "K").Value
        maxTotalVolume = ws.Cells(2, "L").Value
        tickerMaxPercentageIncrease = ws.Cells(2, "I").Value
        tickerMaxPercentageDecrease = ws.Cells(2, "I").Value
        tickerMaxTotalVolume = ws.Cells(2, "I").Value
    
        ' Loop through each row in the data in Column I
        For i = 2 To lastRowColumnI ' Assuming row 1 is the header
            ' Check if the current percentage change is greater than the maximum percentage increase
            If ws.Cells(i, "K").Value > maxPercentageIncrease Then
                maxPercentageIncrease = ws.Cells(i, "K").Value
                tickerMaxPercentageIncrease = ws.Cells(i, "I").Value
            End If
                
            ' Check if the current percentage change is less than the maximum percentage decrease
            If ws.Cells(i, "K").Value < maxPercentageDecrease Then
                maxPercentageDecrease = ws.Cells(i, "K").Value
                tickerMaxPercentageDecrease = ws.Cells(i, "I").Value
            End If
                 
            ' Check if the current total volume is greater than the maximum total volume
            If ws.Cells(i, "L").Value > maxTotalVolume Then
                maxTotalVolume = ws.Cells(i, "L").Value
                tickerMaxTotalVolume = ws.Cells(i, "I").Value
            End If
        Next i
        
        ' Output the results to cells O1 to Q4
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = tickerMaxPercentageIncrease
        ws.Range("Q2").Value = maxPercentageIncrease
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = tickerMaxPercentageDecrease
        ws.Range("Q3").Value = maxPercentageDecrease
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = tickerMaxTotalVolume
        ws.Range("Q4").Value = maxTotalVolume
        
        
        ' Format result to percentage
        ws.Range("Q2").NumberFormat = "0.00%" ' Format to percentage
        ws.Range("Q3").NumberFormat = "0.00%" ' Format to percentage
        
        ' Autofit to display data
        ws.Columns("A:Q").AutoFit

    Next ws
    
    MsgBox ("Tabulation Completed") ' To show that compilation is completed

End Sub



