Attribute VB_Name = "modIniFile"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

' Get the value.
Public Function GetIniString(ByVal Section As String, ByVal SectionKey As String, ByVal iniFileName As String) As String
    On Error GoTo errHandle
    Dim buf As String * 256
    Dim length As Long
    length = GetPrivateProfileString( _
        Section, SectionKey, "<no value>", _
        buf, Len(buf), iniFileName)
    GetIniString = Left$(buf, length)
    Exit Function
errHandle:
End Function

' Set the value.
Public Function SetIniString(ByVal Section As String, ByVal SectionKey As String, ByVal KeyValue As String, ByVal iniFileName As String) As Boolean
    On Error GoTo errHandle
    WritePrivateProfileString _
        Section, SectionKey, _
        KeyValue, iniFileName
    SetIniString = True
    Exit Function
errHandle:
    SetIniString = False
End Function




