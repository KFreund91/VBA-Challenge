Attribute VB_Name = "Module1"
Sub Stocks_Practice()
' Variables
    
    Dim ticker As String
    
    Dim LastRow As Long
    
    Dim year_open As Double
        
    Dim year_close As Double
        
    Dim year_change As Double
        
    Dim percent_change As Double
    
    Dim volume As Long
    
    On Error Resume Next
        
    
'Begin for loop to run through each worksheet

    For Each ws In Worksheets
    
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        ws.Range("K:K").NumberFormat = "0.00%"
'Headers
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        total_volume = 0

        cellnumber = 2
       
        'MsgBox (LastRow)
'For loop
 
 
        
        
        For i = 2 To LastRow
        
        year_open = ws.Cells(i, 3).Value
        year_close = ws.Cells(i, 6).Value
        yearly_change = year_close - year_open
        percent_change = ((year_close - year_open) / year_close) * 100
          
           
            total_volume = total_volume + ws.Cells(i, 7)
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(cellnumber, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(cellnumber, 12).Value = total_volume
                
                ws.Cells(cellnumber, 10).Value = yearly_change
                    'Color conditional
                        If ws.Cells(cellnumber, 10).Value > 0 Then
                        ws.Cells(cellnumber, 10).Interior.ColorIndex = 4
                        Else
                        ws.Cells(cellnumber, 10).Interior.ColorIndex = 3
                        End If
                ws.Cells(cellnumber, 11).Value = percent_change
                
                cellnumber = cellnumber + 1
                total_volume = 0
            
                
                End If
                 
            
                
        Next i
                
        
    Next ws
       
          
End Sub




