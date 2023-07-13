VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Analyze()

For Each ws In Worksheets

Dim ticker As String
Dim column As Integer
column = 1
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim output_row As Integer
output_row = 2

Dim year_open As Double
year_open = ws.Cells(2, 3).Value

Dim year_change As Double
Dim Total_Volume As Double
Total_Volume = 0

Dim percent_change As Double

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

ws.Range("K:K").NumberFormat = "0.00%"


Dim Max As Double
Dim Min As Double
Dim Most_Volume As Double
Dim Max_Index As Integer
Dim Min_Index As Integer
Dim MV_Index As Integer

ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"


For i = 2 To lastrow

    If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
        ticker = ws.Cells(i, 1).Value
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
        year_change = ws.Cells(i, 6).Value - year_open
        
        percent_change = year_change / year_open
        
        
    ws.Range("I" & output_row).Value = ticker
    ws.Range("J" & output_row).Value = year_change
            If year_change < 0 Then
            ws.Range("J" & output_row).Interior.ColorIndex = 3
            ElseIf year_change > 0 Then
            ws.Range("J" & output_row).Interior.ColorIndex = 4
            End If
    ws.Range("l" & output_row).Value = Total_Volume
    ws.Range("K" & output_row).Value = percent_change
        
        output_row = output_row + 1
        year_open = ws.Cells(i + 1, 3).Value
        Total_Volume = 0
        
        Else
        
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
        End If
     Next i
     
    ws.Range("P2").Value = Application.WorksheetFunction.Max(Columns("K"))
        ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").Value = Application.WorksheetFunction.Min(Columns("K"))
        ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("P4").Value = Application.WorksheetFunction.Max(Columns("L"))
     

Max = Application.WorksheetFunction.Max(Columns("K"))
Min = Application.WorksheetFunction.Min(Columns("K"))
Most_Volume = Application.WorksheetFunction.Max(Columns("L"))
Max_Index = WorksheetFunction.Match(Max, Range("K:K"), 0)
Min_Index = WorksheetFunction.Match(Min, Range("K:K"), 0)
MV_Index = WorksheetFunction.Match(Most_Volume, Range("L:L"), 0)
ws.Range("O2").Value = ws.Cells(Max_Index, 9).Value
ws.Range("O3").Value = ws.Cells(Min_Index, 9).Value
ws.Range("O4").Value = ws.Cells(MV_Index, 9).Value
     
            
Next ws


End Sub

