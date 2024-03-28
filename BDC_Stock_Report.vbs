VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ReadData()

For Each ws In ThisWorkbook.Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Vol"
ws.Cells(2, 15).Value = "Greatest % Inc"
ws.Cells(3, 15).Value = "Greatest % Dec"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

Dim LastRow As Long
Dim i As Long

LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

Dim SummaryTableRow As Long
SummaryTableRow = 2

Dim FirstDOYPrice As Double
Dim LastDOYPrice As Double
Dim Change As Double
Dim PercentChange As Double

'Borrowed from CreditCardChecker-CellComparison class example
Dim StockTotal As Double
StockTotal = 0

For i = 2 To LastRow

    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    
        ws.Cells(SummaryTableRow, 9).Value = ws.Cells(i, 1).Value
        
        FirstDOYPrice = ws.Cells(i, 3).Value
        
        SummaryTableRow = SummaryTableRow + 1
        
        StockTotal = ws.Cells(i, 7).Value
     
     
    Else
    
        'Borrowed from CreditCardChecker-CellComparison class example
        StockTotal = StockTotal + ws.Cells(i, 7).Value

    
    End If
    
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    ws.Cells(SummaryTableRow - 1, 12).Value = StockTotal
    
    'Found formatting examples on Stack Overflow
    ws.Cells(SummaryTableRow - 1, 12).NumberFormat = "#,##"
    
    LastDOYPrice = ws.Cells(i, 6).Value
    
    Change = LastDOYPrice - FirstDOYPrice
    
    ws.Cells(SummaryTableRow - 1, 10) = Change
    
    PercentChange = Change / FirstDOYPrice
    
    ws.Cells(SummaryTableRow - 1, 11) = PercentChange
    
    'Found formatting examples on Stack Overflow
    ws.Cells(SummaryTableRow - 1, 11).NumberFormat = "0.00%"
    
        If Change >= 0 Then
        
        ws.Cells(SummaryTableRow - 1, 10).Interior.Color = RGB(0, 255, 0)
        
        Else
        
        ws.Cells(SummaryTableRow - 1, 10).Interior.Color = RGB(255, 0, 0)
        
        End If
    
    End If
            
Next i

Dim GreatestIncTickerID As String
Dim GreatestDecTickerID As String
Dim GreatestVolTickerID As String

GreatestInc = WorksheetFunction.Max(ws.Range("K:K"))

GreatestDec = WorksheetFunction.Min(ws.Range("K:K"))

GreatestVol = WorksheetFunction.Max(ws.Range("L:L"))

ws.Range("Q2").Value = GreatestInc
ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("Q3").Value = GreatestDec
ws.Range("Q3").NumberFormat = "0.00%"

ws.Range("Q4").Value = GreatestVol
ws.Range("Q4").NumberFormat = "#,##"
            
'Found example of WorksheetFunction.Index on scales.arabpyschology.com/stats
GreatestIncTickerID = WorksheetFunction.Index(ws.Range("I:I"), _
WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K:K"), 0))
            
GreatestDecTickerID = WorksheetFunction.Index(ws.Range("I:I"), _
WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K:K"), 0))
            
GreatestVolTickerID = WorksheetFunction.Index(ws.Range("I:I"), _
WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L:L"), 0))
                
ws.Range("P2").Value = GreatestIncTickerID
ws.Range("P3").Value = GreatestDecTickerID
ws.Range("P4").Value = GreatestVolTickerID
        
'Found formatting examples on Stack Overflow
With ws.Range("I1:L1")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
ws.Columns("A:S").AutoFit

Next ws


End Sub



