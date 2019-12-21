# ExcelScripts
Collection of Useful Excel VBA scripts


""" 
Script to split multiple worksheets in a workbook into separate Excel files

Credit to https://superuser.com/users/239751/hrvoj3e for this wonderful solution found in https://superuser.com/questions/561923/how-can-one-split-an-excel-xlsx-file-that-contains-multiple-sheets-into-sep

Sub CreateNewWBS()
Dim wbThis As Workbook
Dim wbNew As Workbook
Dim ws As Worksheet
Dim strFilename As String

    Set wbThis = ThisWorkbook
    For Each ws In wbThis.Worksheets
        strFilename = wbThis.Path & "/" & ws.Name
        ws.Copy
        Set wbNew = ActiveWorkbook
        wbNew.SaveAs strFilename
        wbNew.Close
    Next ws
End Sub

"""
