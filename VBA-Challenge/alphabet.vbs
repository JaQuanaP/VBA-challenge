Attribute VB_Name = "Module1"
Sub StocksandOutputs():

' declare a vsriable to hold the row count
Dim rowCount As Integer

' variable for the totalvolume
Dim totalvolume As Double
totalvolume = 0 ' start the volume total at 0
' variable to keep track for the stock changes for the summary data (Colums I and J)
Dim summaryRow  As Integer
summaryRow = 2 ' starts on row 2 of columns I

'variable

 ' use xlUp command to get the last row / count of rows
 rowCount = Cells(Rows.Count, "A") .End (xlUp) .Row
 
 ' Loop through Column A and check to see where we have changes
 For Row = 2 To rowCount
 
 'simply track changes
 If Cells (Row, 1) . Value <> Cells(Row = 1, 1) .Value
 'display changed ticker in Column G
 
 'display the yearly change in Column J
 Else
 ' if the ticker does not change, simply
 ' add on to the total of the ticker
 totalvolume = volumetoal + Cells(Row, 3).Value

 
 End If
 Next Row
 
 
 
 
 Next Row
 



End Sub


