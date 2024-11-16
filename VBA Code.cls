VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub loopthr()
    Dim cntr As Long
    Dim i As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim curr As String
    Dim prev_curr As String
    Dim ncntr As Long
    ncntr = 2
    Dim ocntr As Long
    ocntr = 2
    Dim stock_volume As LongLong
    total_volume = 0
    Dim current_volume As LongLong
    current_volume = 0
    Dim lastrow_2 As Long
    
    
    prev_curr = Cells(2, 1).Value
   
    
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
    
        total_volume = total_volume + Cells(i, 7)
        current_volume = Cells(i, 7)
        
        curr = Cells(i, 1)
        ocntr = ocntr + 1
        
        If i = 2 Then
            open_price = Cells(2, 3)
            
        End If
        
        If curr <> prev_curr Then
            Cells(ncntr, 9).Value = prev_curr
            
            If ncntr <> 2 Then
                open_price = Cells(i - ocntr, 3).Value
            End If
            
            close_price = Cells(i - 1, 6).Value
            
            
            
            Cells(ncntr, 10).Value = close_price - open_price
            Cells(ncntr, 11).Value = (close_price - open_price) / open_price
            
            Cells(ncntr, 12).Value = total_volume - current_volume
            total_volume = current_volume
            
            
            prev_curr = curr
            
            ncntr = ncntr + 1
            ocntr = 0
            
        End If
        
    Next i
    Cells(ncntr, 9).Value = Cells(lastRow, 1).Value
    
    open_price = Cells(lastRow - ocntr, 3).Value
    close_price = Cells(lastRow, 6).Value
    
    Cells(ncntr, 10).Value = close_price - open_price
    Cells(ncntr, 11).Value = (close_price - open_price) / open_price
    Cells(ncntr, 12).Value = total_volume
    
    Range("K2:K" & lastRow).NumberFormat = "0.00%"
    
    lastrow_2 = Cells(Rows.Count, 12).End(xlUp).Row
    
    minvalue = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_2))
    
    minRow = Application.WorksheetFunction.Match(minvalue, Range("K2:K" & lastrow_2), 0) + 1
    
    Cells(3, 17).Value = Cells(minRow, 11)
    Range("Q3").NumberFormat = "0.00%"
    Cells(3, 16).Value = Cells(minRow, 9)
    
    maxvalue = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_2))
    
    maxRow = Application.WorksheetFunction.Match(maxvalue, Range("K2:K" & lastrow_2), 0) + 1
    
    
    Cells(2, 17).Value = Cells(maxRow, 11)
    Range("Q2").NumberFormat = "0.00%"
    Cells(2, 16).Value = Cells(maxRow, 9)
    
    greatestVolume = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_2))
    greatestRow = Application.WorksheetFunction.Match(greatestVolume, Range("L2:L" & lastrow_2), 0) + 1
    
    Cells(4, 17) = greatestVolume
    
    Cells(4, 16) = Cells(greatestRow, 9).Value
    
       
    
End Sub

