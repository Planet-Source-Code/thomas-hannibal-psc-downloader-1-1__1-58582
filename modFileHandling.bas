Attribute VB_Name = "modFileHandling"
Option Explicit

'***************************************************************************
'Kopiere en fil fra Source til Desitnation
'***************************************************************************
'
'Returnerer filens størrelse, hvis kopieringen lykkes ellers nul (0)
'
'***************************************************************************

Public Function CopyFile(src As String, dst As String) As Single
    'L. Serflaten 1996
    Static buf As String 'Buf$
    Dim BTest As Single 'BTest!
    Dim FSize As Single 'FSize!
    Dim Chunk As Integer    'Chunk%
    Dim F1 As Integer   'F1%
    Dim F2 As Integer   'F2%
    Const BUFSIZE = 2048
    '
    'A larger BUFSIZE is best, but do not attempt to exceed
    '64 K (60000 would be fine)
    '
    If Dir(src) = "" Then MsgBox "File not found": Exit Function

    If Len(Dir(dst)) Then Exit Function
        'If MsgBox(UCase(dst) & Chr(13) & Chr(10) & "File exists. Overwrite?", 4) = 7 Then Exit Function
        'Kill dst
    'End If

    '
    On Error GoTo FileCopyError
    F1 = FreeFile
    Open src For Binary As F1
    F2 = FreeFile
    Open dst For Binary As F2
    
    FSize = LOF(F1)
    BTest = FSize - LOF(F2)
    
    'frmProgress.Show
    
    Do
        If BTest < BUFSIZE Then
            Chunk = BTest
        Else
            Chunk = BUFSIZE
        End If
        buf = String(Chunk, " ")
        Get F1, , buf
        Put F2, , buf
        BTest = FSize - LOF(F2)
        '
        ' __Call percent display here__
        'PercentDone.Caption = ( 100 - Int(100 * BTest/FSize) )
        'PercentDone.Refresh
        '
        'frmProgress.ProgressBar1.Value = (100 - Int(100 * BTest / FSize))
    Loop Until BTest = 0


    Close F1
    Close F2
    CopyFile = FSize
    'Unload frmProgress
    Exit Function
    '
FileCopyError:
    MsgBox "Copy Error!"
    Close F1
    Close F2
End Function

'***********************************************************************************
' Returnerer atributter på fil i flere formater
'***********************************************************************************

Public Function GetFileAttributes(ThisFileName As String, Lan As String, Short As Boolean) As String
CurSub = "modFileHandling -> GetFileAttributes"
On Error GoTo ErrHandle
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    'Nedenstående linie fucker det vist op!!!!
    'Set f = fs.GetFile(fs.GetFileName(ThisFileName))
    
    'Modificeret linie
    Set f = fs.GetFile(ThisFileName)
    
    Select Case f.Attributes
    Case 0
        If Short Then
            GetFileAttributes = "none"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Normal fil - ingen atributter sat!"
        Else
            GetFileAttributes = "Normal file - no atributes are set!"
        End If
        End If
    Case 1
        If Short Then
            GetFileAttributes = "+R"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Kun læsbar"
        Else
            GetFileAttributes = "Read-Only"
        End If
        End If
    Case 2
        If Short Then
            GetFileAttributes = "+H"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Skjult fil"
        Else
            GetFileAttributes = "hidden file"
        End If
        End If
    Case 3
        If Short Then
            GetFileAttributes = "+R+H"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Kun læsbar skjult fil"
        Else
            GetFileAttributes = "Read-Only + hidden file"
        End If
        End If
    Case 4
        If Short Then
            GetFileAttributes = "+S"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "System fil"
        Else
            GetFileAttributes = "System file"
        End If
        End If
    Case 5
        If Short Then
            GetFileAttributes = "+R+S"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Kun læsbar system fil"
        Else
            GetFileAttributes = "Read-Only system file"
        End If
        End If
    Case 6
        If Short Then
            GetFileAttributes = "+H+S"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Skjult system fil"
        Else
            GetFileAttributes = "Hidden system file"
        End If
        End If
    Case 7
        If Short Then
            GetFileAttributes = "+R+H+S"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Kun læsbar og skjult system fil"
        Else
            GetFileAttributes = "Read-Only and hidden system file"
        End If
        End If
    Case 8
        If Short Then
            GetFileAttributes = "Volume"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Disk drev navn"
        Else
            GetFileAttributes = "Disk drive volume label"
        End If
        End If
    Case 16
        If Short Then
            GetFileAttributes = "Folder"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Mappe"
        Else
            GetFileAttributes = "Folder or directory"
        End If
        End If
    Case 32
        If Short Then
            GetFileAttributes = "Archive"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Arkiv fil - Er ændret siden sidste backup"
        Else
            GetFileAttributes = "Archive - File has changed since last backup"
        End If
        End If
    Case 33
        If Short Then
            GetFileAttributes = "+R+Archive"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Kun læsbar Arkiv fil - Er ændret siden sidste backup"
        Else
            GetFileAttributes = "Read-Only Archive - File has changed since last backup"
        End If
        End If
    Case 34
        If Short Then
            GetFileAttributes = "+H+Archive"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Skjult arkiv fil - Er ændret siden sidste backup"
        Else
            GetFileAttributes = "Hidden Archive - File has changed since last backup"
        End If
        End If
    Case 64
        If Short Then
            GetFileAttributes = "Alias"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Alias - Genvej"
        Else
            GetFileAttributes = "Alias - Link or shortcut"
        End If
        End If
    Case 128
        If Short Then
            GetFileAttributes = "Compressed"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Komprimeret fil"
        Else
            GetFileAttributes = "Compressed file"
        End If
        End If
    Case Else
        If Short Then
            GetFileAttributes = "N/A"
        Else
        If Lan = "DK" Then
            GetFileAttributes = "Ukendt!"
        Else
            GetFileAttributes = "Unknown!"
        End If
        End If
    End Select
Exit Function
    
ErrHandle:
Dim MyDescription As String

'Angiver fejlmeddelelse efter fejlnummer
Select Case Err.Number
    Case 53
        GetFileAttributes = "No file found"
        Exit Function
    Case Else
        MyDescription = "Der opstod en fejl under læsningen af fil-attributter."
End Select

ShowErr Err.Number, MyDescription, Err.Description, CurSub ', frmSetup2
End Function


