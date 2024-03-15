Attribute VB_Name = "Module2"
Sub Additional_analysis()

   Dim maxinc As Double
   Dim maxdec As Double
   Dim maxtv As Double
   Dim ws As Worksheet
   
  For Each ws In Worksheets
   
   ws.Range("O2").Value = "Greatest % Increase"
   ws.Range("O3").Value = "Greatest % Decrease"
   ws.Range("O4").Value = "Greatest Total Volume"
   ws.Range("P1").Value = "Ticker"
   ws.Range("Q1").Value = "Value"
   
   tlastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
   maxinc = ws.Cells(2, 11).Value
   maxdec = ws.Cells(2, 11).Value
   maxtv = ws.Cells(2, 12).Value
      
   For j = 3 To tlastrow
      If (ws.Cells(j, 11).Value > maxinc) Then
       maxinc = ws.Cells(j, 11).Value
       ws.Range("Q2").Value = maxinc
       ws.Range("P2").Value = ws.Cells(j, 9).Value
      ElseIf (ws.Cells(j, 11).Value < maxdec) Then
       maxdec = ws.Cells(j, 11).Value
       ws.Range("Q3").Value = maxdec
       ws.Range("P3").Value = ws.Cells(j, 9).Value
      End If
      
      If (ws.Cells(j, 12).Value > maxtv) Then
       maxtv = ws.Cells(j, 12).Value
       ws.Range("Q4").Value = maxtv
       ws.Range("P4").Value = ws.Cells(j, 9).Value
      End If
      
   Next j
   
   ws.Range("Q2:Q3").NumberFormat = "0.00%"
   
  Next ws

End Sub
