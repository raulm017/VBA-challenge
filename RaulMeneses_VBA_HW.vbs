Sub VBAHW_STOCK()
Dim ticker As String
Dim open_value, close_value, percent_change, volume, difference As Double
Dim start As Long
'Loop through worksheets
For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


volume = 0
start = 2
j = 2


last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To last_row

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

volume = volume + ws.Cells(i, 7).Value
ticker = ws.Cells(i, 1)
ws.Cells(j, 9).Value = ticker
ws.Cells(j, 12).Value = volume

open_value = ws.Cells(start, 3).Value
close_value = ws.Cells(i, 6).Value

difference = close_value - open_value

ws.Cells(j, 10).Value = difference

If open_value = 0 Then
open_value = 1
End If
                

ws.Cells(j, 11).Value = difference / open_value '*****

'conditional formatting that will highlight positive change in green and negative change in red
If ws.Cells(j, 10).Value >= 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 4
Else
ws.Cells(j, 10).Interior.ColorIndex = 3


End If


start = i + 1
volume = 0
j = j + 1

Else
volume = volume + ws.Cells(i, 7).Value

End If

Next i

Next ws

End Sub
