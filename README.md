# VBA--challenge
week 2 challenge

complete code creation for assignment.
Uploaded code as a bas file; 
Module2 VBA challenge codeFINAL

extensiverun time for full Multiple_year_stock_data file
I started running the code at 10pm on 1/9 and it is still running as of 8am 1/10.  assingment without screenshots

full code:
Sub VBA_challenge()

For Each ws In Worksheets

'step one
Dim ticker As String
Dim volumn As Double

Dim Summary_Table_Row As Integer

'step two
Dim openprice As Variant
Dim closeprice As Variant
Dim pricechange As Variant
Dim percentchange As Variant
'step 3
Dim greatp As Variant
Dim lowp As Variant
Dim greatv As Variant
Dim tgreatp As String
Dim tlowp As String
Dim tgreatv As String


Summary_Table_Row = 2

volumn = 0

closeprice = 0
pricechange = 0
pecentchange = 0

greatp = 0
lowp = 0
greatv = 0
tgreatp = 0
tlowp = 0
tgreatv = 0
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  'nameing the table
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volumn"
ws.Cells(1, 14) = "Greatest % increase"
ws.Cells(2, 14) = "Greatest % decrease"
ws.Cells(3, 14) = "Greatest total volume"

'creating the LOOP

For i = 2 To LastRow

'identifing the tickers and total volumn

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        ticker = ws.Cells(i, 1).Value
        volumn = volumn + ws.Cells(i, 7).Value
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("L" & Summary_Table_Row).Value = volumn
       Summary_Table_Row = Summary_Table_Row + 1
        'volumn reset
         volumn = 0
    Else
        volumn = volumn + ws.Cells(i, 7).Value
End If

' stock end of year price change

     If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        openprice = ws.Cells(i, 3).Value
       
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i + 2, 1) Then
        closeprice = ws.Cells(i + 1, 6).Value
      End If
         
    pricechange = closeprice - openprice
    ws.Range("J" & Summary_Table_Row).Value = pricechange
    
    percentchange = (pricechange / openprice)
    ws.Range("K" & Summary_Table_Row).Value = percentchange
    ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
                
'changes color for positive or negative

If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
Else
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
End If

Next i


LastRow2 = Cells(Rows.Count, 11).End(xlUp).Row

For j = 2 To LastRow2
'created additional functions
' greatest and lowest percent change
If greatp < ws.Cells(j, 11).Value Then
    greatp = ws.Cells(j, 11).Value
    tgreatp = ws.Cells(j, 9).Value
End If

If lowp > ws.Cells(j, 11).Value Then
    lowp = ws.Cells(j, 11).Value
    tlowp = ws.Cells(j, 9).Value
End If

'greatest volumn
If greatv < ws.Cells(j, 12).Value Then
    greatv = ws.Cells(j, 12).Value
    tgreatv = ws.Cells(j, 9).Value
End If

'filling the table

ws.Cells(1, 15) = tgreatp
ws.Cells(2, 15) = tlowp
ws.Cells(3, 15) = tgreatv

ws.Cells(1, 16) = greatp
ws.Cells(2, 16) = lowp
ws.Cells(3, 16) = greatv
ws.Range("P1:P2").NumberFormat = "0.00%"

Next j

Range("J:J,K:K,L:L,N:N,P:P").EntireColumn.AutoFit

Next ws
    
End Sub

