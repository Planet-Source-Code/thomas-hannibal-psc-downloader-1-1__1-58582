VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "PSC-Downloader"
   ClientHeight    =   4935
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   2625
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   2625
   Begin VB.TextBox TxtDDE 
      Height          =   525
      Left            =   2040
      TabIndex        =   13
      Text            =   "TxtDDE"
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin PSCdl.LaVolpeButton cmd 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   900
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "LaVolpeButton1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form1.frx":1272
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2115
      Top             =   3870
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   2115
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":128E
      Top             =   4365
      Width           =   420
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1410
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2340
      Width           =   1950
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   8
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Picture2 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   9
      Top             =   0
      Width           =   0
   End
   Begin PSCdl.LaVolpeButton cmd 
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "LaVolpeButton1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form1.frx":1294
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin PSCdl.LaVolpeButton cmd 
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "LaVolpeButton1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form1.frx":12B0
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin PSCdl.LaVolpeButton cmd 
      Height          =   285
      Index           =   3
      Left            =   0
      TabIndex        =   12
      Top             =   1680
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "LaVolpeButton1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form1.frx":12CC
      ALIGN           =   1
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
      OPTVAL          =   0   'False
      OPTMOD          =   0   'False
      GStart          =   0
      GStop           =   16711680
      GStyle          =   0
   End
   Begin VB.Image imgClosedOK 
      Height          =   480
      Left            =   1260
      Picture         =   "Form1.frx":12E8
      Top             =   4365
      Width           =   480
   End
   Begin VB.Image imgOpenPutIn 
      Height          =   480
      Left            =   1215
      OLEDropMode     =   1  'Manual
      Picture         =   "Form1.frx":1BB2
      Top             =   3960
      Width           =   480
   End
   Begin VB.Image imgUp 
      Height          =   240
      Left            =   1710
      Picture         =   "Form1.frx":247C
      Top             =   4230
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   240
      Left            =   1710
      Picture         =   "Form1.frx":2A06
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   1125
      Picture         =   "Form1.frx":2F90
      Top             =   585
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      OLEDropMode     =   1  'Manual
      Picture         =   "Form1.frx":351A
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FileName"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   3
      Left            =   45
      TabIndex        =   5
      Top             =   585
      Width           =   645
   End
   Begin VB.Image imgSucces 
      Height          =   480
      Left            =   675
      Picture         =   "Form1.frx":3DE4
      Top             =   4365
      Width           =   480
   End
   Begin VB.Image imgError 
      Height          =   480
      Left            =   90
      Picture         =   "Form1.frx":46AE
      Top             =   4365
      Width           =   480
   End
   Begin VB.Image imgOpen 
      Height          =   480
      Left            =   630
      OLEDropMode     =   1  'Manual
      Picture         =   "Form1.frx":4F78
      Top             =   3915
      Width           =   480
   End
   Begin VB.Image imgClosed 
      Height          =   480
      Left            =   45
      Picture         =   "Form1.frx":5842
      Top             =   3915
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Log"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   2115
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   2
      Left            =   0
      Top             =   2025
      Width           =   2130
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   1
      Left            =   -45
      Top             =   810
      Width           =   2130
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   80
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Zip"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   2
      Left            =   2070
      TabIndex        =   2
      Top             =   585
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Image"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   1
      Left            =   1395
      TabIndex        =   1
      Top             =   585
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      ForeColor       =   &H8000000E&
      Height          =   195
      Index           =   0
      Left            =   675
      TabIndex        =   0
      Top             =   585
      Width           =   510
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   2070
      Picture         =   "Form1.frx":610C
      Top             =   90
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1410
      Picture         =   "Form1.frx":69D6
      Top             =   90
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   750
      Picture         =   "Form1.frx":72A0
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public MenuVisible As Boolean
Private MeActive As Boolean
Private MoveMe As Boolean
Private showPath As Boolean
Private showLog As Boolean
Private iX As Integer
Private iY As Integer
Private dPath
Private fName As String
Private img As Integer
Private mOver As Boolean
Dim sMenu As Integer
Dim img1OK As Boolean
Dim img2OK As Boolean
Dim img3OK As Boolean
Dim img4OK As Boolean

Dim timg As Integer




Private Sub cmd_Click(Index As Integer)
    Select Case Index
    Case 0
        Form2.Show 1, Me
    Case 1
    Dim OldPath As String
    Dim newPath As String
        If dPath = "" Then
            OldPath = App.Path
        Else
            OldPath = dPath
        End If
        newPath = BrowseForFolder(Me.hWnd, "The destination folder is currently set to: " & OldPath & ". Choose your new destination folder!", OldPath)
        If newPath <> "" Then dPath = newPath
        Text2.Text = Text2.Text & "Dest.path: " & dPath & vbCrLf
        cmd(1).ToolTipText = dPath

        'Debug.Print dPath
    Case 2
        If showLog Then
            showLog = False
            cmd(2).Caption = "Show log"
            Me.Height = 2020
        Else
            showLog = True
            cmd(2).Caption = "Hide log"
            Me.Height = 3800
            'Timer1.Enabled = False
        End If
        'Call Resize
    Case 3
        FileSave Text2.Text, App.Path & "\dl.log"
        Unload Me
        End
    End Select
End Sub


Private Sub Form_Deactivate()
    'If MenuVisible Then
    '    Me.Height = Shape1(1).Top + Shape1(1).Height
    '    MenuVisible = False
    'End If
End Sub

Private Sub Form_GotFocus()
    'If MenuVisible Then
    '    Me.Height = Shape1(1).Top + Shape1(1).Height - 10
    '    MenuVisible = False
    'End If
End Sub

Private Sub Form_Load()
    'Load Form2
    cmd(0).Top = Shape1(1).Top + Shape1(1).Height - 10
    cmd(0).Left = 0
    cmd(0).Width = Me.Width
    cmd(1).Top = cmd(0).Top + cmd(0).Height
    cmd(1).Left = 0
    cmd(1).Width = Me.Width
    cmd(2).Top = cmd(1).Top + cmd(1).Height
    cmd(2).Left = 0
    cmd(2).Width = Me.Width
    cmd(3).Top = cmd(2).Top + cmd(2).Height
    cmd(3).Left = 0
    cmd(3).Width = Me.Width
    
    'setting caption
    cmd(0).Caption = "HowTo + About"
    cmd(1).Caption = "Destination path"
    cmd(2).Caption = "Show log screen"
    cmd(3).Caption = "Exit"
    
    Shape1(2).Top = cmd(3).Top + cmd(3).Height
    'Text2.Top = Shape1(2).Top + Shape1(2).Height + 20
    'Text2.Left = 20
    'Text2.Width = Me.Width - cmd2.Width - 100
    'cmd2.Top = Shape1(2).Top + Shape1(2).Height + 20
    'cmd2.Left = Text2.Left + Text2.Width + 50
    
    'Shape1(3).Top = cmd2.Top + cmd2.Height
    
    Label2.Top = Shape1(2).Top + Shape1(2).Height + 20
    Label2.Left = 20
    Text2.Top = Label2.Top + Label2.Height
    Text2.Left = 20
    Text2.Width = Me.Width - 40
    
    'Timer
    Timer1.Interval = 500
    timg = 0
    
    'Initial form settings
    Me.Left = Val(GetIniString("POSITION", "Left", App.Path & "\" & App.EXEName & ".ini"))
    Me.Top = Val(GetIniString("POSITION", "Top", App.Path & "\" & App.EXEName & ".ini"))
    dPath = GetIniString("PATHS", "Destination", App.Path & "\" & App.EXEName & ".ini")
    cmd(1).ToolTipText = dPath
    
    Me.BackColor = RGB(50, 110, 160)
    Text2.BackColor = Me.BackColor
    

    Me.Height = Shape1(1).Top + Shape1(1).Height - 10
    putMeOnTop Me
    Text2.Text = Format(Now, "dd-mmm-yyyy") & vbCrLf & "Dest.path=" & dPath & vbCrLf
    
    sMenu = 0
    Dim t As Integer
    For t = 0 To 2
        Shape1(t).Left = 0
        Shape1(t).Width = Me.Width
    Next t

End Sub

Private Sub Form_LostFocus()
    'If MenuVisible Then
    '    Me.Height = Shape1(1).Top + Shape1(1).Height
    '    MenuVisible = False
    'End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = True
    iX = x
    iY = y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = False
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim MoveX As Integer
    Dim MoveY As Integer
    
    If MoveMe Then
        MoveX = x - iX
        MoveY = y - iY
        'Debug.Print MoveX
        'Debug.Print MoveY
        Me.Left = Me.Left + MoveX
        Me.Top = Me.Top + MoveY
        'Debug.Print X & " - " & Y
    End If
    'Debug.Print y

End Sub



Private Sub Form_Unload(Cancel As Integer)
    Debug.Print dPath
    SetIniString "POSITION", "Left", Me.Left, App.Path & "\" & App.EXEName & ".ini"
    SetIniString "POSITION", "Top", Me.Top, App.Path & "\" & App.EXEName & ".ini"
    SetIniString "PATHS", "Destination", dPath, App.Path & "\" & App.EXEName & ".ini"
End Sub

Private Sub Image1_DblClick()
    fName = ""
    Set Image1.Picture = imgOpenPutIn.Picture
    Set Image2.Picture = imgClosed.Picture
    Set Image3.Picture = imgClosed.Picture
    Set Image4.Picture = imgClosed.Picture
    Image2.OLEDropMode = 0
    Image3.OLEDropMode = 0
    Image4.OLEDropMode = 0
    img1OK = False
    img2OK = False
    img3OK = False
    img4OK = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = True
    iX = x
    iY = y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveFrm x, y
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = False
End Sub

Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'if an url/text is drag and dropped from browser
    Dim s As String
    Dim sUrl As String
    Dim sPageTitle As String
    
    If dPath = "" Then
      MsgBox "No destination path available..."
      Exit Sub
    End If
    If Data.GetFormat(vbCFText) Then
        'fName = Trim(Data.GetData(vbCFText))
        s = Trim(Data.GetData(vbCFText))
        fName = ExtTrim(s, False, True, "_")
        'Debug.Print fName
        While Asc(Right(fName, 1)) < 40
            fName = Left(fName, Len(fName) - 1)
        Wend
        Debug.Print fName 'Asc(Right(fName, 1))
        Set Image1.Picture = imgClosedOK.Picture
        Set Image2.Picture = imgOpenPutIn.Picture
        Set Image3.Picture = imgOpenPutIn.Picture
        Set Image4.Picture = imgOpenPutIn.Picture
        Image2.OLEDropMode = 1
        Image3.OLEDropMode = 1
        Image4.OLEDropMode = 1
        img1OK = True

        Text2.Text = Text2.Text & vbCrLf & "<" & fName & ">"
        
        'Creating a new directory with the source code name as foldername!
        If Len(Dir(dPath & "\" & fName, vbDirectory)) = 0 Then
          MkDir (dPath & "\" & fName)
        End If
    
        ' Get url from browser and make shortcut in directory
        GetURLfromBrowser sUrl, sPageTitle
        If Len(sUrl) <> 0 Then
          CreateHyperlink dPath & "\" & fName & "\" & fName & "_at_PSC.url", sUrl
        End If
    End If
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = True
    iX = x
    iY = y
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveFrm x, y
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = False
End Sub

Private Sub Image2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'if an url/text is drag and dropped from browser
    Dim r As Boolean
    Dim f As String
    If dPath = "" Then
      MsgBox "No destination path available..."
      Exit Sub
    End If
    If Data.GetFormat(vbCFText) Then
        img = 2
        'Timer2.Enabled = True
        Text1.Text = Data.GetData(vbCFText)
        Set Image2.Picture = imgClosed.Picture
        r = FileSave(Text1.Text, dPath & "\" & fName & "\" & fName & "_about.txt")
        
        'Timer2.Enabled = False
        If r Then
            Set Image2.Picture = imgSucces.Picture
        Else
            Set Image2.Picture = imgError.Picture
        End If
        
        f = fName & "_about.txt"
        If r Then
            f = f & vbCrLf & "> Succes!"
        Else
            f = f & vbCrLf & ">>>> Failure!"
        End If

        Text2.Text = Text2.Text & vbCrLf & " " & f
        img2OK = True
        If img1OK And img2OK And img3OK And img4OK Then Image1_DblClick
    End If
End Sub


Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = True
    iX = x
    iY = y
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveFrm x, y
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = False
End Sub

Private Sub Image3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If dPath = "" Then
      MsgBox "No destination path available..."
      Exit Sub
    End If
    If Data.GetFormat(vbCFFiles) Then
        Dim sfile As String
        Dim fn As String
        Dim i As Integer
        Dim r As Boolean
        Dim numFiles As Integer
        Dim f As String
        img = 3
        'Timer2.Enabled = True
        numFiles = Data.Files.Count
        For i = 1 To numFiles
            If (GetAttr(Data.Files(i)) And vbDirectory) <> vbDirectory Then
                sfile = Data.Files(i) 'Skip directory and get the first file
                Debug.Print sfile
                Set Image3.Picture = imgClosed.Picture
                Exit For
            Else
                sfile = Data.Files(i) 'Skip directory and get the first file
                Debug.Print sfile
                Set Image3.Picture = imgClosed.Picture
                Exit For
            End If
        Next i
        If fName = "" Then
            fn = sfile
        Else
            fn = fName & Right(sfile, 4)
        End If

        r = DL(sfile, dPath & "\" & fName & "\" & fn)
        'Timer2.Enabled = False
        
        If Not FileReal(dPath & "\" & fName & "\" & fn) Then
          CopyFile sfile, dPath & "\" & fName & "\" & fn
          Debug.Print sfile
        End If
        
        If r Then
            Set Image3.Picture = imgSucces.Picture
        Else
            Set Image3.Picture = imgError.Picture
        End If
        f = fn
        If r Then
            f = f & vbCrLf & "> Succes!"
        Else
            f = f & vbCrLf & ">>>> Failure!"
        End If
        
        Text2.Text = Text2.Text & vbCrLf & " " & f
        img3OK = True
        If img1OK And img2OK And img3OK And img4OK Then Image1_DblClick
        'Call frmMainFileOpenAs(sFile) 'Open the file
    End If
End Sub

Private Sub Image4_Click()
    'timg = 0
    'Timer1.Enabled = True
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = True
    iX = x
    iY = y
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveFrm x, y
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = False
End Sub

Private Sub Image4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'if an url/text is drag and dropped from browser
    If dPath = "" Then
      MsgBox "No destination path available..."
      Exit Sub
    End If
    Dim sfile As String
    Dim fn As String
    Dim r As Boolean
    Dim f As String
    If Data.GetFormat(vbCFText) Then
        img = 4
        'Timer1.Enabled = True
        'Timer2.Enabled = True
        
        sfile = Data.GetData(vbCFText)
        Set Image4.Picture = imgClosed.Picture
        
        If fName = "" Then
            fn = sfile
        Else
            fn = fName & ".zip"
        End If
        r = DL(sfile, dPath & "\" & fName & "\" & fn)
        
        'Timer1.Enabled = False
        'Timer2.Enabled = False
        If r Then
            Set Image4.Picture = imgSucces.Picture
        Else
            Set Image4.Picture = imgError.Picture
        End If
        
        f = fn
        If r Then
            f = f & vbCrLf & "> Succes!"
        Else
            f = f & vbCrLf & ">>>> Failure!"
        End If

        Text2.Text = Text2.Text & vbCrLf & " " & f
        
        img4OK = True
        If img1OK And img2OK And img3OK And img4OK Then Image1_DblClick

    End If
End Sub

Private Sub Image5_Click()
    If sMenu = 0 Then
        If showLog Then
            Me.Height = 3800
        Else
            Me.Height = 2020
        End If
        Set Image5.Picture = imgUp.Picture
        sMenu = 1
    Else
        Me.Height = Shape1(1).Top + Shape1(1).Height - 10
        Set Image5.Picture = imgDown.Picture
        sMenu = 0
    End If
End Sub



Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = True
    iX = x
    iY = y
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveFrm x, y
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    MoveMe = False
End Sub




Private Sub Timer1_Timer()
    Select Case timg
    Case 0
        Set Image4.Picture = imgOpen.Picture
        timg = 1
    Case 1
        Set Image4.Picture = imgOpenPutIn.Picture
        timg = 2
    Case 2
        Set Image4.Picture = imgClosed.Picture
        timg = 0
    End Select
    Debug.Print timg
End Sub


Private Sub MoveFrm(ByVal x As Integer, ByVal y As Integer)
    Dim MoveX As Integer
    Dim MoveY As Integer
    
    If MoveMe Then
        MoveX = x - iX
        MoveY = y - iY
        'Debug.Print MoveX
        'Debug.Print MoveY
        Me.Left = Me.Left + MoveX
        Me.Top = Me.Top + MoveY
        'Debug.Print X & " - " & Y
    End If
End Sub
