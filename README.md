# VBA-Challenge
Excel-VBA Scripting Module Two Challenge 


VBA-Module 2 Code

Sub StockDataSet()
'Set Dimensions

    Dim Total As Double
    Dim RowIndex As Long
    Dim Change As Double
    Dim ColumnIndex As Integer
    Dim STart As Long
    Dim RowCount As Long
    Dim PercentChange As Double
    Dim Days As Integer
    Dim DailyChange As Single
    Dim AvgChange As Double
    Dim ws As Worksheet
    
'-------------------------------------------------------------------

For Each ws In Worksheets
 ColumnIndex = 0
 Total = 0
 Change = 0
 STart = 2
 DailyChange = 0
 'Set Title Row
 
 '------------------------------------
 
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Volume"
 ws.Range("P1").Value = "Ticker "
 ws.Range("Q1").Value = "Value"
 ws.Range("Q2").Value = "Greatest % Increase"
 ws.Range("Q3").Value = "Greatest % Decrease"
 ws.Range("Q4").Value = "Greatest Total Volume"
 
 'Get the Row Number of the Last row of data
 
'--------------------------------------------------
RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

    For RowIndex = 2 To RowCount
       If ws.Cells(RowIndex + 1, 1).Value <> ws.Cells(RowIndex, 1).Value Then
        Total = Total + ws.Cells(RowIndex, 7).Value
        
        If Total = 0 Then
        ws.Range("I" & 2 + ColumnIndex).Value = Cells(RowIndex, 1).Value
        ws.Range("J" & 2 + ColumnIndex).Value = 0
        ws.Range("k" & 2 + ColumnIndex).Value = "%" & 0
        ws.Range("L" & 2 + ColumnIndex).Value = 0
        
            Else
            
            If ws.Cells(STart, 3) = 0 Then
                For find_value = STart To RowIndex
                    If ws.Cells(find_value, 3).Value <> 0 Then
                    STart = find_value
                    Exit For
                    End If
            Next find_value
            End If
            
            Change = (ws.Cells(RowIndex, 6) - ws.Cells(STart, 3))
            
            PercentChange = Change / ws.Cells(STart, 3)
            
            STart = RowIndex + 1
            
            ws.Range("I" & 2 + ColumnIndex) = ws.Cells(RowIndex, 1).Value
            ws.Range("J" & 2 + ColumnIndex) = Change
            ws.Range("j" & 2 + ColumnIndex).NumberFormat = "0.00"
            ws.Range("K" & 2 + ColumnIndex).Value = PercentChange
            ws.Range("K" & 2 + ColumnIndex).NumberFormat = "0.00"
            ws.Range("l" & 2 + ColumnIndex).Value = Total
            
             
            Select Case Change
            Case Is > 0
                ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 4
            Case Is < 0
                ws.Range("j" & 2 + ColumnIndex).Interior.ColorIndex = 3
            Case Else
                ws.Range("j" & 2 + ColumnIndex).Interior.ColorIndex = 0
            End Select
             
        End If
        
        Total = 0
        Change = 0
        ColumnIndex = ColumnIndex + 1
        Days = 0
        DailyChange = 0
        
            Else
            Total = Total + ws.Cells(RowIndex, 7).Value
             
       End If
    
    Next RowIndex
ws.Range("q2") = "%" & WorksheetFunction.Max(ws.Range("K2:k" & RowCount)) * 100
ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("l2:k" & RowCount)) * 100
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))


increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("k2:k" & RowCount)), ws.Range("K2:k" & RowCount), 0)
decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:k" & RowCount)), ws.Range("K2:k" & RowCount), 0)
volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:l" & RowCount)), ws.Range("L2:L" & RowCount), 0)

ws.Range("P2") = ws.Cells(increase_number + 1, 9)
ws.Range("p3") = ws.Cells(decrease_number + 1, 9)
ws.Range("p4") = ws.Cells(volume_number + 1, 9)

Next ws
  
End Sub

