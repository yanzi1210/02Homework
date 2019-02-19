# 02Homework
Sub Uniqueticker()
Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
WS.Activate
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Change"
Cells(1, 12).Value = "Total Stock Volumn"

  Dim Row As Long
  Row = 2
  
  Dim total As LongLong
  total = 0
  
  Dim Ticker As String
  Ticker = Cells(2, 1).Value
  
  Dim Open_Price As Double
  'Set Initial Open Price
  Open_Price = Cells(2, 3).Value
  
  Dim Close_Price As Double
  
  Dim Yealy_Change As Double
  
  Dim Percent_Change As Double
  
  
  lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To lastrow
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        
        'Set Close Price
        Close_Price = Cells(i, 6).Value
        Yearly_Change = Close_Price - Open_Price
        
        total = total + Cells(i, 7).Value
        
        Range("I" & Row).Value = Ticker
        
        Range("L" & Row).Value = total
        
        Range("J" & Row).Value = Yearly_Change
      
        'Add Percent Change
        If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
        Else
                Percent_Change = Yearly_Change / Open_Price
                Range("K" & Row).Value = Percent_Change
                Range("K" & Row).NumberFormat = "0.00%"
        End If
        Row = Row + 1
        Open_Price = Cells(i + 1, 3)
        total = 0
   
    Else
    total = total + Cells(i, 7).Value
    End If
    Next i
    
    NLastRow = WS.Cells(Rows.Count, 10).End(xlUp).Row
    For n = 2 To NLastRow
     If (Cells(n, 10).Value >= 0) Then
                Cells(n, 10).Interior.ColorIndex = 4
            ElseIf Cells(n, 10).Value < 0 Then
                Cells(n, 10).Interior.ColorIndex = 3
            End If
        Next n
        
    Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        ' Look through each rows to find the greatest value and its associate ticker
        For h = 2 To NLastRow
            If Cells(h, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & NLastRow)) Then
                Cells(2, 16).Value = Cells(h, 9).Value
                Cells(2, 17).Value = Cells(h, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(h, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & NLastRow)) Then
                Cells(3, 16).Value = Cells(h, 9).Value
                Cells(3, 17).Value = Cells(h, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(h, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & NLastRow)) Then
                Cells(4, 16).Value = Cells(h, 9).Value
                Cells(4, 17).Value = Cells(h, 12).Value
            End If
        Next h
    Next WS
End Sub
   
