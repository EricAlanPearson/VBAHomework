Sub moderate():
 For Each ws In Worksheets
 Dim i As Long
 Dim j As Long
 Dim totalv As Double
 Dim ticker As String
 Dim lastrow As Long
 Dim ychng As Double
 Dim pchng As Double
 Dim oprice As Double
 Dim cprice As Double
 Dim oprice_row As Long
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Value"
 totalv = 0
 j = 2
 oprice_row = 2
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 For i = 2 To lastrow
     If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then
         totalv = totalv + ws.Range("G" & i).Value
     Else
         ticker = ws.Range("A" & i).Value
         oprice = ws.Range("C" & oprice_row)
         cprice = ws.Range("F" & i)
         ychng = cprice - oprice
         If oprice = 0 Then
            pchng = 0
         Else
            pchng = ychng / oprice
         End If
         ws.Range("I" & j).Value = ticker
         ws.Range("L" & j).Value = totalv + ws.Range("G" & i).Value
         ws.Range("J" & j).Value = ychng
         ws.Range("K" & j).Value = pchng
         ws.Range("K" & j).NumberFormat = "0.00%"
         If ws.Range("J" & j).Value > 0 Then
            ws.Range("J" & j).Interior.ColorIndex = 4
         Else
            ws.Range("J" & j).Interior.ColorIndex = 3
         End If
         j = j + 1
         totalv = 0
         oprice_row = i + 1
     End If
 Next i
 Next ws
End Sub
