Attribute VB_Name = "Module1"
Sub Stock_analysis()
 
  Dim ticker As String
  Dim openingprice As Double
  Dim closingprice As Double
  Dim yearlychange As Double
  Dim totalstockvolume As Double
  Dim tablerow As Double
  Dim openingrow As Double
  Dim ws As Worksheet
  
 For Each ws In Worksheets
  
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percentage Change"
  ws.Range("L1").Value = "Total Stock Volume"
  
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  tablerow = 2
  openingrow = 2
  
    For i = 2 To lastrow
      
      If (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
      
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
      
      Else
        
        ticker = ws.Cells(i, 1).Value
        ws.Cells(tablerow, 9).Value = ticker
        
        openingprice = ws.Cells(openingrow, 3).Value
        closingprice = ws.Cells(i, 6).Value
        yearlychange = closingprice - openingprice
        ws.Cells(tablerow, 10).Value = yearlychange
          If (yearlychange < 0) Then
           ws.Cells(tablerow, 10).Interior.Color = RGB(255, 0, 0)
          Else
           ws.Cells(tablerow, 10).Interior.Color = RGB(0, 255, 0)
          End If
        
        ws.Cells(tablerow, 11).Value = yearlychange / openingprice
      
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
        ws.Cells(tablerow, 12).Value = totalstockvolume
        
        tablerow = tablerow + 1
        openingrow = i + 1
        totalstockvolume = 0
        
      End If
          
    Next i
    
    ws.Range("K:K").NumberFormat = "0.00%"
  
  Next ws
           
End Sub
