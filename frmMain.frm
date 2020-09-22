VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Access Password Finder"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   1950
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkUse3 
      Caption         =   "Try other Key"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3840
      TabIndex        =   8
      Top             =   990
      Width           =   1170
   End
   Begin VB.CommandButton btnPass2000 
      Caption         =   "Find 2000 Password"
      Height          =   375
      Left            =   2925
      TabIndex        =   6
      Top             =   615
      Width           =   2055
   End
   Begin VB.CommandButton btnPass97 
      Caption         =   "Find 97 Password"
      Height          =   375
      Left            =   1125
      TabIndex        =   5
      Top             =   615
      Width           =   1815
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1125
      TabIndex        =   2
      Top             =   1335
      Width           =   2175
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "..."
      Height          =   375
      Left            =   4845
      TabIndex        =   1
      Top             =   255
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1275
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1125
      TabIndex        =   0
      Top             =   255
      Width           =   3735
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "Written by Dan Fogelberg 1999"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3585
      TabIndex        =   7
      Top             =   1725
      Width           =   1845
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1140
      TabIndex        =   4
      Top             =   1095
      Width           =   855
   End
   Begin VB.Label lblFileName 
      Caption         =   "File Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1125
      TabIndex        =   3
      Top             =   15
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------
'This is a handy little routine for retrieving
'forgotten passwords out of Access97 and 2000.
'
'The original code for Access97 was found on a
'newsgroup and I translated it to VB.  It would
'not work on 2000 so I fixed that!  The 97 code
'should also work on Access 2+, but I have not
'checked that.
'
'Do not abuse this.  This should be used only for
'databases you originally had permissions to and
'forgot the password.
'
'-Dan Fogelberg
'DanFogelberg@newmail.net
'
'Disclaimer:
'Use this code at your own risk!  I take no responsibility
'for it.
'-------------------------------------------------------
'Use:
'   Select the appropriate MDB file and then
'   Click the button for the version.  If you
'   Do not know the version try both... :-)
'   If you get gibberish well then it must be
'   the other!  Enjoy and send comments to above.

Private Sub btnPass2000_Click()
    On Error GoTo errHandler
    Dim ch(40) As Byte
    Dim x As Integer, sec2, intChar As Integer, blnUse3 As Boolean
    If Trim(txtFileName) = "" Then Exit Sub
    'Used integers instead of hex :-)  Easier to read
    sec2 = Array(0, 194, 117, 236, 55, 25, 202, 156, 250, 130, 208, 40, 230, 87, 56, 138, 96, 16, 26, 123, 54, 177, 252, 223, 177, 51, 122, 19, 67, 139, 33, 177, 51, 112, 239, 121, 91, 214, 59, 124, 42)
    'I found that some DB's use this scheme see below for the logic to determine which is which :-)
    sec3 = Array(0, 229, 117, 236, 55, 62, 202, 156, 250, 165, 208, 40, 230, 112, 56, 138, 96, 55, 26, 123, 54, 150, 252, 223, 177, 20, 122, 19, 67, 172, 33, 177, 51, 87, 239, 121, 91, 241, 59, 124, 42)
    txtPass.Text = ""
    blnUse3 = False
    
    Open txtFileName.Text For Binary Access Read As #1 Len = 40
    Get #1, &H42, ch
    Close #1
    'Check to see which key by running through first 6 letters of password
    'This is not foolproof by any means.
    For x = 1 To 6
      intChar = ch(x) Xor sec2(x)
         'This is kind of lame but it assumes that most passwords
         'are in this range of keyboard chars :-)
      If ((intChar < 32) Or (intChar > 126)) And (intChar <> 0) Then
         blnUse3 = True 'Set a flag
      End If
    Next x
    'Allow force of key3.
    If chkUse3.Value = vbChecked Then
      blnUse3 = True
    End If
    'Now solve for password
    For x = 1 To 40
         If blnUse3 = True Then
            intChar = ch(x) Xor sec3(x)
         Else
            intChar = ch(x) Xor sec2(x)
         End If
         txtPass.Text = txtPass.Text & Chr(intChar)
    Next x
    Exit Sub
errHandler:
    MsgBox "ERROR occcured:" & vbCrLf & Err.Number & ":  " & Err.Description, vbCritical, "ERROR"
    Exit Sub
End Sub

Private Sub btnPass97_Click()
On Error GoTo errHandler
    Dim ch(18) As Byte, x As Integer
    Dim sec
    If Trim(txtFileName) = "" Then Exit Sub
    'Used integers instead of hex :-)  Easier to read
    sec = Array(0, 134, 251, 236, 55, 93, 68, 156, 250, 198, 94, 40, 230, 19, 182, 138, 96, 84)
    txtPass.Text = ""
    Open txtFileName.Text For Binary Access Read As #1 Len = 18
    Get #1, &H42, ch
    Close #1
    For x = 1 To 17
        txtPass.Text = txtPass.Text & Chr(ch(x) Xor sec(x))
    Next x
    Exit Sub
errHandler:
    MsgBox "ERROR occcured:" & vbCrLf & Err.Number & ":  " & Err.Description, vbCritical, "ERROR"
    Exit Sub
    
End Sub

Private Sub btnOpen_Click()
On Error GoTo errHandler
    With CommonDialog1
       .CancelError = True 'Causes error on Cancel
       .Filter = "Access (*.mdb)|*.mdb|All Files (*.*)|*.*"
       .FilterIndex = 1
       .ShowOpen
       txtFileName.Text = .FileName
    End With
    Exit Sub
errHandler:
   txtFileName.Text = ""  'Ensure it is empty
   Exit Sub
End Sub


