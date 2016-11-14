Attribute VB_Name = "Module1"
Option Compare Database
Sub XLS_Paste()
    Dim exelName As String
    Dim sheetName As String
    Dim queryName As String
    Dim startCopyCell As String
    Dim endCopyCell As String
    Dim copyArea As String
    

    exelName = "sample.xls"
    sheetName = "Report"
    queryName = "qry_sel_DAILY_DATA"
    startCopyCell = "B10"
    endCopyCell = "L62"
     
    Call XLS_Paste_1(exelName, sheetName, queryName, startCopyCell, endCopyCell)
   

End Sub

Private Sub XLS_Paste_1(exelName As String, sheetName As String, queryName As String, startCopyCell As String, endCopyCell As String)
    'Use Paste to Access standard module

    On Error GoTo Err_XLS_Paste_1

    Dim DB As DAO.Database
    Dim RS As DAO.Recordset
    Dim objApp As Object

    Dim copyArea As String
    
    copyArea = startCopyCell + ":" + endCopyCell
    
    exelName = CurrentProject.path & "\" & exelName
    
    Set DB = CurrentDb
    Set RS = DB.OpenRecordset(queryName)
    
    On Error Resume Next
    
    Set objApp = CreateObject("Excel.Application") 'Excel Object
    
    objApp.Visible = True 'View Excel on the screen
    
    With objApp.Workbooks.Open(exelName)
    
    With objApp.Sheets(sheetName)
        .Range(copyArea).ClearContents 'Clear of Copy Area
        .Range(startCopyCell).CopyFromRecordset RS 'param:startCopyCell standard in output
    
    End With
    
    objApp.Visible = True
    objApp.DisplayAlerts = False
    objApp.Save
    objApp.DisplayAlerts = True
    objApp.Quit
    Set objApp = Nothing
    
    Set RS = Nothing
    Set DB = Nothing
    Set OBJEXE = Nothing
    
    Exit Sub
    
    End With

Exit_XLS_Paste_1:
    Exit Sub

Err_XLS_Paste_1:
    MsgBox Err.Description
    Resume Exit_XLS_Paste_1

End Sub

Private Function Replace_XLS_Sheet_Name_M(oldSheetName As String, createMonth As String) As String
    



End Function


'Get absolute path from relative path
Function GetCurPath2() As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetCurPath2 = FSO.GetAbsolutePathName("")
    Set FSO = Nothing
End Function
