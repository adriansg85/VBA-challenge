

Sub loopStocks()


'Define the variable for checking ticker symbols
Dim firstTicker As String
Dim tickerSybol As String

'Define variables for the excersise
Dim priceChange As Double
Dim percentChange As Double
Dim totalVolume As Double
Dim lastRow As Long
Dim yearClose As Double
Dim yearOpen As Double
Dim summaryRow As Integer

'Sub loop_workbooks_for_loop()

Dim i As Long
Dim ws_num As Integer

Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
ws_num = ThisWorkbook.Worksheets.Count

For j = 1 To ws_num
    ThisWorkbook.Worksheets(j).Activate

summaryRow = 2

totalVolume = 0

yearOpen = ThisWorkbook.Worksheets(j).Cells(2, 3).Value

'Create a Loop to check the document

'Define last row
lastRow = ThisWorkbook.Worksheets(j).Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow

totalVolume = totalVolume + ThisWorkbook.Worksheets(j).Cells(i, 7).Value

If ThisWorkbook.Worksheets(j).Cells(i, 1).Value <> ThisWorkbook.Worksheets(j).Cells(i + 1, 1).Value Then

yearClose = ThisWorkbook.Worksheets(j).Cells(i, 6).Value

priceChange = yearClose - yearOpen

If priceChange > 0 Then

ThisWorkbook.Worksheets(j).Cells(summaryRow, 10).Interior.ColorIndex = 4

Else

ThisWorkbook.Worksheets(j).Cells(summaryRow, 10).Interior.ColorIndex = 3

End If

If yearOpen > 0 Then

percentageChange = priceChange / yearOpen

Else

percentageChange = "NA"

End If

ThisWorkbook.Worksheets(j).Cells(summaryRow, 10) = priceChange
ThisWorkbook.Worksheets(j).Cells(summaryRow, 11) = percentageChange

yearOpen = ThisWorkbook.Worksheets(j).Cells(i + 1, 3).Value

ThisWorkbook.Worksheets(j).Cells(summaryRow, 9) = ThisWorkbook.Worksheets(j).Cells(i, 1).Value

ThisWorkbook.Worksheets(j).Cells(summaryRow, 12) = totalVolume
totalVolume = 0

summaryRow = summaryRow + 1

End If


Next i

Next j


'starting_ws.Activate 'activate the worksheet that was originally active

End Sub






