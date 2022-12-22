Attribute VB_Name = "Module1"
Sub OneYear()
For Each ws In Worksheets
    'Set column titles.
    ws.Range("I1", "P1").Value = "Ticker"
    ws.Range("M1", "O1").Value = " "
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

'Column titles of extra functionality.
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

'Setting variables: yearly change (yrch), percentage change (prch), total stock volume (tsv)
Dim yrch As Double
Dim prch As Double
Dim tsv As Double
Dim currentopen As Double

'Adding functionality: greatest percent increase (grpinc), greatest percent decrease (grpdec), and greatest total volume grtv
Dim grpinc As Double
Dim grpdec As Double
Dim grtv As Double

'Initiating Ticker Sequence and basic loop; finding number of rows, and assigning 'year change'.
Dim numrows As Double
Dim i, j As Double
numrows = ws.Cells(Rows.Count, 1).End(xlUp).Row
currentopen = ws.Cells(2, 3).Value

j = 1
For i = 2 To numrows
    tsv = tsv + ws.Cells(i, 7).Value
    If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
        yrch = ws.Cells(i, 6).Value - currentopen
        prch = yrch / currentopen
        ws.Cells(j + 1, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(j + 1, 10).Value = yrch
        If (yrch < 0) Then
            ws.Cells(j + 1, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(j + 1, 10).Interior.ColorIndex = 4
        End If
        ws.Cells(j + 1, 11).Value = prch
        ws.Cells(j + 1, 12).Value = tsv
        'Here are extra functionality condiitionals concerning greatest % inc, greatest % dec, and greatest tsv
        If (prch > grpinc) Then
            grpinc = prch
            ws.Cells(2, 16).Value = ws.Cells(i, 1).Value
            ws.Cells(2, 17).Value = prch
        ElseIf (prch < grpdec) Then
            grpdec = prch
            ws.Cells(3, 16).Value = ws.Cells(i, 1).Value
            ws.Cells(3, 17).Value = prch
        End If
        If (tsv > grtv) Then
            grtv = tsv
            ws.Cells(4, 16).Value = ws.Cells(i, 1).Value
            ws.Cells(4, 17).Value = tsv
        End If
        j = j + 1
        currentopen = ws.Cells(i + 1, 3).Value
        tsv = 0
    End If
Next i
        
ws.Range("Q2:Q3").NumberFormat = "0.00%"
ws.Range("K:K").NumberFormat = "0.00%"

Next ws

End Sub

Sub numrows()
Dim numrow As Long
numrow = Cells(Rows.Count, 1).End(xlUp).Row
MsgBox ("There are" + Str(numrow) + " rows in this sheet!")
End Sub
