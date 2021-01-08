'Attribute VB_Name = "Step6"



Sub Step6()

'Plan for Step 6 : Move Tab named (Request Number_Surge_IFSM) +  ("Last  Good Pass") + ("Summary Table") into a new Excel Workbook

'TODO:
''****************************************************
'   Move Tab named (Request Number_Surge_IFSM) +  ("Last  Good Pass") + ("Summary Table") into a new Excel Workbook
'   New Excel Workbook is named as (Request Number_Surge_IFSM)

''****************************************************

Dim RqID As Double
RqID = 122133

Dim SheetName1 As String
SheetName1 = RqID & "_Surge_IFSM"

Dim SheetName2 As String
SheetName2 = "Summary Table"

    'Variable declaration
    Dim wb As Workbook
    
    'Create New Workbook
    Set wb = Workbooks.Add
    
    'Save Above Created New Workbook
    Dim relativePath As String
    relativePath = ThisWorkbook.Path & "\" & SheetName1 & ".xls"
    
    wb.SaveAs Filename:=relativePath
    
    
ThisWorkbook.Sheets(SheetName1).Move Before:=wb.Sheets(Sheets.Count)
ThisWorkbook.Sheets(SheetName2).Move Before:=wb.Sheets(Sheets.Count)


wb.Close SaveChanges:=True


End Sub

