Attribute VB_Name = "modSaveTextFile"
Option Explicit

Public Function FileSave(Text As String, FilePath As String) As Boolean
'Save a text file
On Error GoTo error
Dim Directory As String
              Directory$ = FilePath
       On Error GoTo error
       Open Directory$ For Output As #1
           Print #1, Text
       Close #1
       FileSave = True
Exit Function
error:
    'MsgBox Err.Description, vbExclamation, "Error"
    FileSave = False
End Function

