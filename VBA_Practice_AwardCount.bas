Attribute VB_Name = "Module1"
Sub totalStar()
    Worksheets("Sheet1").Activate
    Dim names(2 To 51) As String
    Dim total As Integer
    
    Worksheets("Sheet1").Activate
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Worksheets("Sheet1").Activate
       
            
    For r = 2 To lastRow
        total = 0
        names(r) = Cells(r, 1).Value
        
        If Cells(r, 4).Value = "Full-Star" Then
            total = total + 1
        End If
        
        If Cells(r, 5).Value = "Full-Star" Then
            total = total + 1
        End If
        If Cells(r, 6).Value = "Full-Star" Then
            total = total + 1
        End If
        If Cells(r, 7).Value = "Full-Star" Then
            total = total + 1
        End If
        If Cells(r, 8).Value = "Full-Star" Then
            total = total + 1
        End If
        
        Cells(r, 9).Value = total
       
    Next r
    


End Sub

