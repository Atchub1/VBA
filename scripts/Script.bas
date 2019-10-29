Attribute VB_Name = "Module1"
Sub Stock_data()

Dim Ticker As String
Dim TSV As Double
Dim ws As Worksheet

For Each ws In Worksheets
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"

TSV = 0
'TSV is Total Stock volume

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (lastrow)
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            TSV = TSV + ws.Cells(i, 7).Value
        
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("J" & Summary_Table_Row).Value = TSV
        
            Summary_Table_Row = Summary_Table_Row + 1
        
            TSV = 0
        Else
            TSV = TSV + ws.Cells(i, 7).Value
        End If
        
        Next i

Next




End Sub

