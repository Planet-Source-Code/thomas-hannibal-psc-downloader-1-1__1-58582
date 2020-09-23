VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About - PSC-downloader"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3735
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PSCdl.LaVolpeButton cmd 
      Height          =   285
      Left            =   1260
      TabIndex        =   1
      Top             =   2385
      Width           =   1140
      _ExtentX        =   2011
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
      MICON           =   "Form2.frx":1272
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   2175
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   3525
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const sp As Variant = vbCrLf
Dim clCount As Byte


Private Sub cmd_Click()
    If clCount = 0 Then
        Label1.Caption = "This little program was developed to help my self. It simple helps you to dowbnload source code to a specific directory. It's freeware and OpenSource." & vbCrLf & vbCrLf & "Thomas Braad Hannibal, 2001"
        clCount = 1
        cmd.Caption = "HowTo"
    Else
        Label1.Caption = "1. Mark the project name in your browser" & sp & "2. Drag it to the 'FileName' icon." & sp & "3. Drag description to 'Text' icon." & sp & "4. Drag pic to 'image' icon." & sp & "5. Drag .zip file to 'zip' icon" & sp & "6. Each drag leads to a DL to the destination folder and renaming to the filename." & sp & "- If succesfull the box will turn green else red!" & sp & "7. Db.click on first icon resets the filename! So do dropping items to all icons or dragging a new filename to the filename box!"
        clCount = 0
        cmd.Caption = "About"
    End If
End Sub

Private Sub Form_Load()
    clCount = 0
    Me.Top = Form1.Top
    Me.Left = Form1.Left + Form1.Width + 200
        Label1.Caption = "1. Mark the project name in your browser" & sp & "2. Drag it to the 'FileName' icon." & sp & "3. Drag description to 'Text' icon." & sp & "4. Drag pic to 'image' icon." & sp & "5. Drag .zip file to 'zip' icon" & sp & "6. Each drag leads to a DL to the destination folder and renaming to the filename." & sp & "- If succesfull the box will turn green else red!" & sp & "7. Db.click on first icon resets the filename! So do dropping items to all icons or dragging a new filename to the filename box!"
    cmd.Caption = "About"
    Me.BackColor = Form1.BackColor
End Sub
