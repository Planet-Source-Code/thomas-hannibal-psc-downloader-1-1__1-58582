Attribute VB_Name = "modDLfromInternet"
Option Explicit

Public Function DL(ByVal SourceFile As String, ByVal DestinationFile As String) As Boolean
  Dim obj As clsDownload
  Set obj = New clsDownload
  Dim bRet As Boolean
  
     Screen.MousePointer = vbHourglass
       bRet = obj.Get_File(SourceFile, DestinationFile)
        If bRet = False Then DL = False
          Screen.MousePointer = vbDefault
     Set obj = Nothing
     DL = True
     'MsgBox "Done", vbInformation

End Function

