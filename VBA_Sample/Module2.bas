Attribute VB_Name = "Module2"
Option Compare Database
Const xls1 = "Sample2.xls"
Const xls2 = "Sample1.xls"

Sub Copy_Paste_WorkbookToWorkBookFromAccessVBA()
    
    
    Dim objApp As Object
    Dim xls1FileName As String
    Dim xls2FileName As String
    
    
    Set objApp = CreateObject("Excel.Application")
    
    xls1FileName = CurrentProject.path & "\" & xls1
    xls2FileName = CurrentProject.path & "\" & xls2
    
    'Copy file open
    With objApp.workbooks.Open(xls1FileName)
    
    objApp.Visible = True
    
    'Copy from sheets & copy range
    With .worksheets("Report")
        .range("C5", "F12").copy
    
    End With
    
    End With
    
    'Copy file quit
    objApp.Visible = True
    objApp.displayalerts = False
    objApp.Quit
    
    'Paste file open
    With objApp.workbooks.Open(xls2FileName)
    
    objApp.Visible = True
    
    'Paste to sheets & paste range
    With .worksheets("Report")
        .range("A30").pastespecial Paste:=xlPasteValues
        
    End With
    
    End With
    
    'Paste file save & quit
    objApp.Visible = True
    objApp.displayalerts = False
    objApp.Save
    objApp.displayalerts = True
    objApp.Quit
    Set objApp = Nothing
    
End Sub
