Sub hard():
 For Each ws In Worksheets
 Dim i As Long
 Dim j As Long
 Dim totalv As Double
 Dim ticker As String
 Dim lastrow As Long
 Dim lastrow_Display As Long
 Dim ychng As Double
 Dim pchng As Double
 Dim oprice As Double
 Dim cprice As Double
 Dim oprice_row As Long
 Dim ginc As Double
 Dim gdec As Double
 Dim gtotalv As Double
 Dim ticker_ginc As String
 Dim ticker_gdec As String
 Dim ticker_gtotalv As String
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Value"
 ws.Range("P1").Value = "Ticker"
 ws.Range("Q1").Value = "Value"
 ws.Range("O2").Value = "Greatest % Increase"
 ws.Range("O3").Value = "Greatest % Decrease"
 ws.Range("O4").Value = "Greatest Total Volume"
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
         if oprice = 0 Then
            pchng = 0
         Else
            pchng = ychng / oprice
         End if
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
 ginc = ws.Range("K2" & 2).Value
 gdec = ws.Range("K2" & 2).Value
 gtotalv = ws.Range("L2" & r).Value
 ticker_ginc = ws.Range("I2" & r).Value
 ticker_gdec = ws.Range("I2" & r).Value
 ticker_gtotalv = ws.Range("I2" & r).Value
 lastrow_Display = ws.Cells(Rows.Count, "I").End(xlUp).Row
 For r = 2 to lastrow_Display:
     If ws.Range("K" & r + 1).value > ginc Then
        ginc = ws.Range("K" & r + 1).Value
        ticker_ginc = ws.Range("I" & r +1).Value
     elseif ws.Range("K" & r + 1).value < gdec Then
        gdec = ws.Range("K" & r + 1).Value
        ticker_gdec = ws.Range("I" & r + 1).Value
     elseif ws.Range("L" & r + 1).Value > gtotalv Then
        gtotalv = ws.Range("L" & r + 1).Value
        ticker_gtotalv = ws.Range("I" & r + 1).value
     End If
 Next r
 ws.Range("P2").Value = ticker_ginc
 ws.Range("P3").Value = ticker_gdec
 ws.Range("P4").Value = ticker_gtotalv
 ws.Range("Q2").Value = ginc
 ws.Range("Q3").Value = gdec
 ws.Range("Q4").Value = gtotalv
 ws.Range("Q2:Q3").NumberFormat = "0.00%"
 Next ws
End Sub