Attribute VB_Name = "Module4"
Sub verify()
    Dim i
    Dim j
    Dim k
    Dim l
    Dim m
    Dim n As Integer
    

    
    Dim MaxCol_1
    Dim MaxCol_2
    Dim MaxRow
    
        
  MaxRow = ThisWorkbook.Sheets(1).Cells(1048576, 1).End(xlUp).Row   '最終行
  MaxCol_1 = ThisWorkbook.Sheets("wk").Cells(1, 1).End(xlToRight).Column  'wkの最終列
  MaxCol_2 = ThisWorkbook.Sheets(1).Cells(16, 1).End(xlToRight).Column  'sheet1の最終列

  
  For i = 1 To MaxCol_1
        
    With ThisWorkbook.Sheets("wk")
        .Cells(2, i).Copy
        .Cells(11, i).PasteSpecial
        .Cells(3, i).Copy
        .Cells(14, i).PasteSpecial
        .Cells(4, i).Copy
        .Cells(12, i).PasteSpecial
        .Cells(5, i).Copy
        .Cells(15, i).PasteSpecial
        .Cells(6, i).Copy
        .Cells(13, i).PasteSpecial
    End With
    
 Next i


 
     For n = 1 To MaxCol_1
 
        If ThisWorkbook.Sheets("wk").Cells(8, n) = "1000-3800" Or ThisWorkbook.Sheets("wk").Cells(8, n) = "5100-6000" Then
       
            ThisWorkbook.Sheets("wk").Cells(8, n).Copy
            ThisWorkbook.Sheets("wk").Cells(17, n).PasteSpecial
            
        End If

   
        If ThisWorkbook.Sheets("wk").Cells(8, n) = "1000-3800" Then
            ThisWorkbook.Sheets("wk").Cells(8, n) = "1000-2700"
        ElseIf ThisWorkbook.Sheets("wk").Cells(8, n) = "5100-6000" Then
              ThisWorkbook.Sheets("wk").Cells(8, n) = "5100-5400"
        End If
           
           
        If ThisWorkbook.Sheets("wk").Cells(17, n) = "1000-3800" Then
            ThisWorkbook.Sheets("wk").Cells(17, n) = "3400-3800"
        ElseIf ThisWorkbook.Sheets("wk").Cells(17, n) = "5100-6000" Then
            ThisWorkbook.Sheets("wk").Cells(17, n) = "5800-6000"
     
        End If
                          
          Next n
        
   

    

For j = 1 To MaxCol_1
       For k = 18 To MaxRow
        For l = 7 To MaxCol_2
           
            If ThisWorkbook.Sheets(1).Cells(k, 1) = ThisWorkbook.Sheets("wk").Cells(9, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 2) = ThisWorkbook.Sheets("wk").Cells(10, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 3) = ThisWorkbook.Sheets("wk").Cells(11, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 4) = ThisWorkbook.Sheets("wk").Cells(12, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 5) = ThisWorkbook.Sheets("wk").Cells(13, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 6) = ThisWorkbook.Sheets("wk").Cells(14, j) And _
                ThisWorkbook.Sheets(1).Cells(17, l) = ThisWorkbook.Sheets("wk").Cells(8, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 17) = "" Then

                  ThisWorkbook.Sheets(1).Cells(k, l) = ThisWorkbook.Sheets("wk").Cells(15, j)
            End If
        Next l
    Next k
Next j
  
  For j = 1 To MaxCol_1
    For k = 18 To MaxRow
        For l = 7 To MaxCol_2
            If ThisWorkbook.Sheets(1).Cells(k, 1) = ThisWorkbook.Sheets("wk").Cells(9, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 2) = ThisWorkbook.Sheets("wk").Cells(10, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 3) = ThisWorkbook.Sheets("wk").Cells(11, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 4) = ThisWorkbook.Sheets("wk").Cells(12, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 5) = ThisWorkbook.Sheets("wk").Cells(13, j) And _
                ThisWorkbook.Sheets(1).Cells(k, 6) = ThisWorkbook.Sheets("wk").Cells(14, j) And _
                ThisWorkbook.Sheets(1).Cells(17, l) = ThisWorkbook.Sheets("wk").Cells(17, j) Then
                ThisWorkbook.Sheets(1).Cells(k, l) = ThisWorkbook.Sheets("wk").Cells(15, j)
            End If
            
        Next l
    Next k
  Next j
  

End Sub
