Attribute VB_Name = "Module1"
' Creating a script that loops through all the stocks to calculate price change over one year

Sub StockYearlyChange()

Dim Summary_Table_Row As Integer ' creates placeholder for the table
Summary_Table_Row = 2

Dim Ticker_Name As String ' sets the name of the stock ticker
Ticker_Name = Cells(2, 1).Value

Dim Opening_Price As Double  'sets opening price
Opening_Price = Cells(2, 3).Value

Dim Closing_Price As Double 'sets closing price
Closing_Price = Cells(2, 6).Value

Dim Total_Volume As Double 'Sets total Volume
Total_Volume = 0

Dim Year_Change As Double 'set for total difference of opening and closing stock prices over year
Year_Change = 0

Dim Percent_Change As Double 'set the percentage change for the year's opening/close prices

Dim ws As Worksheet
 
For i = 2 To 753001

'For Each ws In Worksheets

    If Ticker_Name <> Cells(i, 1).Value Then  'compares rows to determine if stock ticker name is different
    
    'Sets the variables to values from the columns
    
    Ticker_Name = Cells(i, 1).Value
    Opening_Price = Cells(i, 3).Value
    Closing_Price = Cells(i, 6).Value
    Total_Volume = Cells(i, 7).Value
    
    'Total price change throughout the year
    Year_Change = Opening_Price - Closing_Price
    
    'Percentage change from opening to closing price for that year (Year_Change/Opening Price) * 100
    Percent_Change = (Year_Change / Opening_Price) * 100
    
    'Total Volume of the shares for each of stock ticker
    Total_Volume = Total_Volume + Cells(i, 7).Value
    
    '-------------------------------------------
    '-------------------------------------------
    
    'Prints the Percentage Change
    Range("K" & Summary_Table_Row).Value = Percent_Change
    
    'Prints the ticker name in the Table
    Range("I" & Summary_Table_Row).Value = Ticker_Name
    
    'Prints the total change price
    Range("J" & Summary_Table_Row).Value = Year_Change
    
    'Prints out total volume traded
    Range("L" & Summary_Table_Row).Value = Total_Volume
    
    'Use conditional formatting to color-coded green for positive changes and red for negative change
    
    If Year_Change > 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        
    ElseIf Year_Change < 0 Then
        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
    
    End If
     
    'Adds one to the new Row and reset the year change
    Summary_Table_Row = Summary_Table_Row + 1
    Year_Change = 0
    Total_Volume = 0
    
   Else
    
   'Total price change throughout the year
    Year_Change = Opening_Price - Closing_Price
    
    'Percentage change from opening to closing price for that year (Year_Change/Opening Price) * 100
    Percent_Change = (Year_Change / Opening_Price) * 100
    
    'Total Volume of the shares for each of stock ticker
    Total_Volume = Total_Volume + Cells(i, 7).Value
    

    End If

    Next i

End Sub
