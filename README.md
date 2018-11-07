1个Excel里200个sheet 怎么将200个sheet合并到一个sheet里？


Sub Macro1()
    Dim dstSheet As Worksheet
    Dim srcSheet As Worksheet
    Dim dstRows As Long
    Dim srcRows As Long
     
    Application.DisplayAlerts = False
     
    Set dstSheet = Sheets(1)
    dstRows = dstSheet.Cells.SpecialCells(xlLastCell).Row
    dstSheet.Activate
     
    While Sheets.Count > 1
        Set srcSheet = Sheets(2)
        srcRows = srcSheet.Cells.SpecialCells(xlLastCell).Row
         
        srcSheet.Rows("1:" & srcRows).Copy
        dstSheet.Range("A" & (dstRows + 1)).Select
        dstSheet.Paste
         
        dstRows = dstRows + srcRows
        srcSheet.Delete
    Wend
     
    Application.DisplayAlerts = True
End Sub

