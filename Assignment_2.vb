Sub AddStockTotals():

    Dim ws As Worksheet
    
'Create For statement to do the same thing to each worksheet
    For Each ws In ActiveWorkbook.Worksheets
        
    'Activate worksheet
        ws.Activate
   
    'Declare variables
        Dim ir, j As Long
        Dim x As Integer
    
        Dim tk As String
        Dim op, cl, vol, per, yr As Double
        
        Dim hd1, hd2, hd3, hd4 As String

        Dim lr, lc As Long
    
    'Count total rows & columns
        lr = Cells(Rows.Count, 1).End(xlUp).Row
        lc = Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Set Headers
        hd1 = "Ticker"
        hd2 = "Yearly Change"
        hd3 = "Percent Change"
        hd4 = "Total Stock Volume"
    
        Cells(1, lc + 2).Value = hd1
        Cells(1, lc + 3).Value = hd2
        Cells(1, lc + 4).Value = hd3
        Cells(1, lc + 5).Value = hd4
        
    'Set base inputs
        ir = 2
        op = Cells(2, 3).Value

    'Set For loop to compare the values in the same column of different rows
        For j = 2 To lr
        
        'Comparing two ticker symbols
            If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
        
            'If the sybmols don't equal, use the values of the first symbol
                tk = Cells(j, 1).Value
                vol = vol + Cells(j, 7).Value
                cl = Cells(j, 6).Value
                yr = cl - op
                
                Cells(ir, lc + 2).Value = tk
                Cells(ir, lc + 5).Value = vol
                Cells(ir, lc + 3).Value = yr
                
            'If statements to avoid dividing by or into 0
                If op = 0 And cl = 0 Then
                
                    per = 0
                    
                ElseIf op = 0 And yr <> 0 Then
                
                    per = 1
                    
                Else
                    
                    per = yr / op
                    
                End If
                
                Cells(ir, lc + 4).Value = per
                
            'Reset values for next symbol
                op = Cells(j + 1, 3).Value
                ir = ir + 1
                vol = 0
        
            Else
                
            'If symbols match, add the volumes
                vol = vol + Cells(j, 7).Value
            
            End If
       
        Next j
    
    'Basic formatting
        Columns(lc + 2).AutoFit
        Columns(lc + 3).AutoFit
        Columns(lc + 4).AutoFit
        Columns(lc + 5).AutoFit
        Columns(lc + 4).NumberFormat = "0.00%"
                
    'Find new "last row" for tables on each worksheet
        Dim lrt As Long
        lrt = Cells(Rows.Count, lc + 2).End(xlUp).Row
        
    'Check if row in table is positive or negative to fill cells
        For x = 2 To lrt
        
            If Cells(x, lc + 4).Value >= 0 Then
            
                Cells(x, lc + 3).Interior.Color = RGB(97, 255, 97)
                
            Else
            
                Cells(x, lc + 3).Interior.Color = RGB(254, 85, 72)
                
            End If
        
        Next x
        
    'Find max values
    
    'Declare new variables
        Dim max, maxu, maxd As Double
        Dim tkM, tkU, tkD As String
        Dim lb1, lb2, lb3 As String
        Dim thd1, thd2 As String
        Dim lct As Integer
        Dim y As Long
        
        lct = Cells(1, Columns.Count).End(xlToLeft).Column
       
    'Set labels for new table
        lb1 = "Greatest Positive Percent Change"
        lb2 = "Greatest Negative Percent Change"
        lb3 = "Greatest Total Volume"
        Cells(2, lct + 2).Value = lb1
        Cells(3, lct + 2).Value = lb2
        Cells(4, lct + 2).Value = lb3
        Columns(lct + 2).AutoFit
        
        thd1 = "Ticker Symbol"
        thd2 = "Values"
        Cells(1, lct + 3).Value = thd1
        Cells(1, lct + 4).Value = thd2
        
    'Set base values
        max = Cells(2, lct).Value
        maxu = Cells(2, lct - 1).Value
        maxd = Cells(2, lct - 1).Value
       
        tkM = Cells(2, lct - 3).Value
        tkU = Cells(2, lct - 3).Value
        tkD = Cells(2, lct - 3).Value
        
    'Compares one a cell in one row to the next
        For y = 2 To lrt
        
        'Finds max volume
            If Cells(y + 1, lct).Value <= max Then
            
                max = max
                tkM = tkM
                
            Else
                
                max = Cells(y + 1, lct).Value
                tkM = Cells(y + 1, lct - 3).Value
                
            End If
            
        'Finds greatest positive change
            If Cells(y + 1, lct - 1).Value <= maxu Then
            
                maxu = maxu
                tkU = tkU
                
            Else
            
                maxu = Cells(y + 1, lct - 1).Value
                tkU = Cells(y + 1, lct - 3).Value
                
            End If
            
        'Find greatest negative change
            If Cells(y + 1, lct - 1).Value >= maxd Then
            
                maxd = maxd
                tkD = tkD
                
            Else
            
                maxd = Cells(y + 1, lct - 1).Value
                tkD = Cells(y + 1, lct - 3).Value
                
            End If
        
        Next y
        
    'Place max values in cells and format
        Cells(2, lct + 3).Value = tkU
        Cells(3, lct + 3).Value = tkD
        Cells(4, lct + 3).Value = tkM
        Cells(2, lct + 4).Value = maxu
        Cells(3, lct + 4).Value = maxd
        Cells(4, lct + 4).Value = max
        Cells(2, lct + 4).NumberFormat = "0.00%"
        Cells(3, lct + 4).NumberFormat = "0.00%"
        Columns(lct + 4).AutoFit
        
        
'Do everything to next worksheet
    Next ws
    
    Worksheets(3).Select
       
End Sub


Sub ClearTestCells():

Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Worksheets
        
        ws.Activate
   
        Columns(14).Clear
        Columns(15).Clear
        Columns(16).Clear
        Columns(17).Clear
        
        Range("I1:P291").ClearFormats
                
    Next ws
    
    Worksheets(3).Select
    
End Sub


