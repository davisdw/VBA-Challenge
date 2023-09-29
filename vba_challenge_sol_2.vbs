Attribute VB_Name = "Module2"
' Creating an sub routine to loop through the percent change to determine the highest and lowest percent changes and the highest volume
' may need to use WorkSheetFunction to loop through the percentage changes and Total Volume Columns
' next; determine how to read through the first iteration for highest and lowest percentage changes
' Determine which of the value is the highest and lowest percentage change
        


'Dim Find_Max As Double
'Dim Find_Min As Double
'Dim Find_Max_Vol As Long


    ' loop through the "K" and "L" Columns and compare each value
    ' find and locate the greatest/least percentage change an highest volume traded
    ' once located those values assign them to variables above
    ' in turn have those values display on the specified cells

'-----------------------------------------------------------------------------

Sub FindingGreatestValues()

Dim Find_Greatest_Value  As Double
Dim Find_Least_Value As Double
Dim Find_Greatest_Volume As Long '**
Dim Show_Ticker_Name As String

Dim tickerName1 As String
tickerName1 = "RKS"

Dim tickerName2 As String
tickerName2 = "SGU"

Dim tickerName3 As String
tickerName3 = "PNOW"

' ** Found that the Highest amount of Total  Volume cannot able to hold Long integer

With Worksheets("2018")

Find_Greatest_Value = Application.WorksheetFunction.Max(.Range("K2:K3000"))
Find_Least_Value = Application.WorksheetFunction.Min(.Range("K2:K3000"))
Find_Greatest_Volume = Application.WorksheetFunction.Min(.Range("L2:L3000"))

For j = 2 To 3000

If Cells(j, 9).Value = tickerName1 Then
Cells(4, 17).Value = tickerName1

ElseIf Cells(j, 9).Value = tickerName2 Then
Cells(5, 17).Value = tickerName2

ElseIf Cells(j, 9).Value = tickerName3 Then
Cells(6, 17).Value = tickerName3


End If

Next j

'Find a way to add the ticket name associated with the values

Cells(4, 18).Value = Find_Greatest_Value
Cells(5, 18).Value = Find_Least_Value
Cells(6, 18).Value = Find_Greatest_Volume

End With
    
End Sub

