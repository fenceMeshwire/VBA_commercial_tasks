Option Explicit

Sub accumulate_total_costs()

Dim intQty As Integer
Dim lngRow, lngRowMax As Long
Dim curUnitPrice, curSubtotal, curTotal As Currency

' Data structure: 
' A       B               C
' qty     description     unit price
' 1       article_name    $24.50
' 2       ...             ...

With Sheet1

  lngRowMax = .Cells(.Rows.Count, 1).End(xlUp).Row
  
  For lngRow = 2 To lngRowMax
  
    intQty = .Cells(lngRow, 1).Value
    curUnitPrice = .Cells(lngRow, 3).Value
    
    curSubtotal = intQty * curUnitPrice
    curTotal = curTotal + curSubtotal
    
  Next lngRow

End With

' Output of the total amount:
Debug.Print curTotal

End Sub
