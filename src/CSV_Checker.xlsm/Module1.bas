Attribute VB_Name = "Module1"
Sub delete()
    
    Dim lastrow As Long
    
   lastrow = Cells(Rows.Count, 1).End(xlUp).Row
   
    
    Range(Cells(16, 5), Cells(lastrow, 5)).delete shift:=xlShiftToLeft
   
    Worksheets("sheet1").Range("A16:O" & lastrow).Copy
    
    Sheets.Add after:=ActiveSheet
    ActiveSheet.Paste
    
    
End Sub
