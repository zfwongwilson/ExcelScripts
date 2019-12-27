# ExcelScripts
Collection of Useful Excel VBA scripts

 
### VBA Script to split multiple worksheets in a workbook into separate Excel files ###

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

### Excel formula to extract age from Malaysian NRIC ###

    =119-LEFT(CELL,2)
  
119 refers to the current year 2019. Change accordingly based on the year of calculation. (e.g. for year 2020 it would be 120)

### Excel formula to extract state code from Malaysian NRIC ###

    =MID(CELL,8,2)
  
For NRIC without dashes, use 7 instead of 8.

