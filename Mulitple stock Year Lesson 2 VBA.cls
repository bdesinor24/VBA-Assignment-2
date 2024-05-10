VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Analysis_week_2()
    Dim ticker_name As String
    Dim Summary_Table_ROW As Double
    Summary_Table_ROW = 2
    Dim Initial_open As Double
    Initial_open = Cells(2, 3)
    Dim Initial_Close As Double
    Initial_Close = 1
    Dim Quarterly_Change As Double
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Quarterly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Dim i As Double
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For Each ws In Worksheets
     
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1) <> Cells(i, 1) Then
            ws.Cells(Summary_Table_ROW, 9) = Cells(i, 1)
           'Cells(Summary_Table_ROW, 9) = Cells(i, 3)'
            Initial_Close = Cells(i, 6)
            Quarterly_Change = Initial_Close - Initial_open
            ws.Cells(Summary_Table_ROW, 10) = Quarterly_Change
            ws.Cells(Summary_Table_ROW, 11) = (Quarterly_Change / Initial_open)
            Initial_open = Cells(i + 1, 3)
            Volume = Volume + Cells(i, 7)
            ws.Cells(Summary_Table_ROW, 12) = Volume
            Volume = 0
            Select Case Quarterly_Change
                    Case Is > 0
                        Cells(Summary_Table_ROW, 10).Interior.ColorIndex = 4
                        Case Is < 0
                        Cells(Summary_Table_ROW, 10).Interior.ColorIndex = 3
            End Select
                        Summary_Table_ROW = Summary_Table_ROW + 1
                        
        Else: Volume = Volume + Cells(i, 7)
        
        End If
    Next i
   
   Range("Q2") = WorksheetFunction.Max(Range("J2:J1502"))
   Range("P2") = Cells(WorksheetFunction.Match(WorksheetFunction.Max(Range("J2:J" & lastrow)), Range("J2:J" & lastrow), 0) + 1, 8).Value
   Range("Q3") = WorksheetFunction.Min(Range("J2:J1502"))
   Range("P3") = Cells(WorksheetFunction.Match(WorksheetFunction.Min(Range("J2:J" & lastrow)), Range("J2:J" & lastrow), 0) + 1, 8).Value
   Range("Q4") = WorksheetFunction.Max(Range("K2:K1502"))
   Range("P4") = Cells(WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0) + 1, 8).Value
Next ws
End Sub

            
            
                        
            

            
            
