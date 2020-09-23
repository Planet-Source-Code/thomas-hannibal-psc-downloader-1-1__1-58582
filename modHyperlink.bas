Attribute VB_Name = "modHyperlink"
Option Explicit

' CreateHyperlink is made by: Unknown
' All his submissions at PSC: Unknown
'
' PSC Project Title:  Create a Internet Shorcut on Anyone's Computer
' PSC Project Url: http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=12248&lngWId=1
'
' Description: This code will create an internet shortcut on someone's computer. All you have to do
' it call it with a path and hyperlink!
'
' The code is slitly changed by Thomas Hannibal

Public Function CreateHyperlink(Path As String, Hyperlink As String) As Boolean
  CreateHyperlink = False
    On Error GoTo errHandle
    
    'check if extention .url is added to path... Else add it!
    If Right(Path, 3) <> "url" And Mid(Path, Len(Path) - 4, 1) <> "." Then
      Path = Path & ".url"
    End If
    
    Open Path For Output As #1 'open file access
    Print #1, "[Internetshortcut]" 'print On first line
    Print #1, "URL=" & Hyperlink 'print url On second line
    Close #1 'close it
    CreateHyperlink = True
    Exit Function
errHandle:
    CreateHyperlink = False
    Exit Function
End Function
