Attribute VB_Name = "modGetBrowserUrl"
Option Explicit

' GetURLfromBrowser is made from code by Mark E.
' All his submissions at PSC: http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=331878042&strAuthorName=Mark%20E.&txtMaxNumberOfEntriesPerPage=25)
'
' PSC Project Title: Get displayed url from browser using DDE
' PSC Project Url: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=6808&lngWId=1
'
' Description: This code retrieves the url and window title of any open instance of Netscape Navigator or Internet Explorer.
' It uses DDE (Dynamic data exchange) and WWW_GetWindowInfo and is much more reliable than the FindWindow Function.
' DDE is supported in Internet Explorer > 3.0 and in Netscape Navigator > 2.0.
'
' The code is slitly changed by Thomas Hannibal

Public Function GetURLfromBrowser(url As String, title As String)

 On Error GoTo GUBErrHandler
    Form1.TxtDDE.LinkTopic = ActiveBrowser
    
    ' tell ie to send us
    ' name and title of the last active
    ' window or frame
    
    Form1.TxtDDE.LinkItem = &HFFFFFFFF
    Form1.TxtDDE.LinkMode = 2
    Form1.TxtDDE.LinkRequest
    
    ' parse out info given to us by ie in
    ' form1.TxtDDE.Text; should be in the form
    '        "URL","Page title"
    
    Dim cc As Long, parms(2) As String, quoting As Boolean
    Dim thisParm As Integer, p As Long, c As Byte
    Dim i As Integer
    
    thisParm = 1
    quoting = False
    For i = 1 To Len(Form1.TxtDDE)
        c = Asc(Mid(Form1.TxtDDE, i, 1))
        Select Case c
            Case 34     ' quotation mark
                quoting = Not quoting
            Case 44     ' comma
                If Not quoting Then
                    thisParm = thisParm + 1
                    If thisParm > 2 Then Exit For
                End If
            Case Else
                If quoting Then
                    parms(thisParm) = parms(thisParm) & Chr(c)
                End If
        End Select
    Next i
    
    url = parms(1)
    title = parms(2)
    Exit Function
    
GUBErrHandler:
    ' skip process if any errors occur, i.e., Netscape
    ' did not respond to DDE initiate event
    MsgBox "Browser not loaded."
    On Error GoTo 0
End Function


' This part is made by Thomas Hannibal (ie. Me...)
' For use with GetURLfromBrowser()

Private Function ActiveBrowser() As String
  'Testing for known browsers...
  On Error GoTo errHandle
  
    'Testing for Internet Explorer...
    Form1.TxtDDE.LinkTopic = "iexplore|WWW_GetWindowInfo"
    Form1.TxtDDE.LinkItem = &HFFFFFFFF
    Form1.TxtDDE.LinkMode = 2
    ActiveBrowser = "iexplore|WWW_GetWindowInfo"
    Exit Function

    'Testing for Mozilla Firefox...
    Form1.TxtDDE.LinkTopic = "Firefox|WWW_GetWindowInfo"
    Form1.TxtDDE.LinkItem = &HFFFFFFFF
    Form1.TxtDDE.LinkMode = 2
    ActiveBrowser = "Firefox|WWW_GetWindowInfo"
    Exit Function

    'Testing for Internet Explorer...
    Form1.TxtDDE.LinkTopic = "NETSCAPE|WWW_GetWindowInfo"
    Form1.TxtDDE.LinkItem = &HFFFFFFFF
    Form1.TxtDDE.LinkMode = 2
    ActiveBrowser = "NETSCAPE|WWW_GetWindowInfo"
    Exit Function

    'Testing for Internet Explorer...
    Form1.TxtDDE.LinkTopic = "Mozilla|WWW_GetWindowInfo"
    Form1.TxtDDE.LinkItem = &HFFFFFFFF
    Form1.TxtDDE.LinkMode = 2
    ActiveBrowser = "Mozilla|WWW_GetWindowInfo"
    Exit Function
  
  
  
  Exit Function
errHandle:
  Select Case Err.Number
    Case 282
      'if connection to known browser isn't responding then try the next known browser...
      Resume Next
    Case Else
      Exit Function
    End Select
End Function
