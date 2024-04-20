VERSION 5.00
Begin VB.Form frmExitType 
   BackColor       =   &H003E472C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmExit 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H008080FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00EAF1F4&
         Caption         =   "EXIT PROGRAM"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4200
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CommandButton cmdSwitchUser 
         BackColor       =   &H00EAF1F4&
         Caption         =   "SWITCH USER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   1695
         Left            =   4320
         Picture         =   "frmExitType.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   2415
         Left            =   720
         Picture         =   "frmExitType.frx":201CC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmExitType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public reply As String

Private Sub cmdCancel_Click()
If frmMain.cmdLogOut.Visible = True Then
    frmMain.Show
    Unload Me
    frmMain.cmdLogOut.BackColor = &H82D1B0
    frmMain.Enabled = True
End If

If frmMain.cmdManageAccounts.Caption = "EXIT" Then
    frmMain.Show
    Unload Me
    frmMain.cmdManageAccounts.BackColor = &H82D1B0
    frmMain.Enabled = True
End If
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If frmMain.cmdLogOut.Visible = True Then
        frmMain.Show
        Unload Me
        frmMain.cmdLogOut.BackColor = &H82D1B0
        frmMain.Enabled = True
    End If
    
    If frmMain.cmdManageAccounts.Caption = "EXIT" Then
        frmMain.Show
        Unload Me
        frmMain.cmdManageAccounts.BackColor = &H82D1B0
        frmMain.Enabled = True
    End If
End If
End Sub

Private Sub cmdExit_Click()

reply = MsgBox("Exit Program?", vbInformation + vbYesNo, "")

If reply = vbYes Then
    'DELETING PREVIOUS RECORDS
    frmMain.printsalesdata.RecordSource = "select * from tblprintsales"
    frmMain.printsalesdata.Refresh
    
    If Not frmMain.printsalesdata.Recordset.EOF Then
        Do
        frmMain.printsalesdata.Recordset.Delete
        frmMain.printsalesdata.Refresh
        Loop Until frmMain.printsalesdata.Recordset.EOF
    End If
    
    frmMain.cmdSort.Caption = "SORT"
    frmMain.cmdSort.BackColor = &H82D1B0
    
    If Not frmMain.transactdata.Recordset.EOF Then
        Do
        frmMain.transactdata.Recordset.Delete
        frmMain.transactdata.Refresh
        Loop Until frmMain.transactdata.Recordset.EOF
    End If
    
    End
End If
End Sub

Private Sub cmdExit_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If frmMain.cmdLogOut.Visible = True Then
        frmMain.Show
        Unload Me
        frmMain.cmdLogOut.BackColor = &H82D1B0
        frmMain.Enabled = True
    End If
    
    If frmMain.cmdManageAccounts.Caption = "EXIT" Then
        frmMain.Show
        Unload Me
        frmMain.cmdManageAccounts.BackColor = &H82D1B0
        frmMain.Enabled = True
    End If
End If
End Sub

Private Sub cmdSwitchUser_Click()

reply = MsgBox("Switch Account?", vbInformation + vbYesNo, "")

If reply = vbYes Then
    'DELETING PREVIOUS RECORDS
    frmMain.printsalesdata.RecordSource = "select * from tblprintsales"
    frmMain.printsalesdata.Refresh
    
    If Not frmMain.printsalesdata.Recordset.EOF Then
        Do
        frmMain.printsalesdata.Recordset.Delete
        frmMain.printsalesdata.Refresh
        Loop Until frmMain.printsalesdata.Recordset.EOF
    End If
    
    frmMain.cmdSort.Caption = "SORT"
    frmMain.cmdSort.BackColor = &H82D1B0
    
    frmMain.transactdata.RecordSource = "select * from tbltransact"
    frmMain.transactdata.Refresh
    
    If Not frmMain.transactdata.Recordset.EOF Then
        Do
        frmMain.transactdata.Recordset.Delete
        frmMain.transactdata.Refresh
        Loop Until frmMain.transactdata.Recordset.EOF
    End If
    
    frmLogin.Show
    Unload frmMain
    Unload Me
    frmLogin.txtPassword1.Text = ""
    frmLogin.txtPassword2.Text = ""
    frmLogin.txtUsername1.Text = ""
    frmLogin.txtUsername2.Text = ""
    frmLogin.frmUserType.Visible = True
    frmLogin.frmadminlogin.Visible = False
    frmLogin.frmuserlogin.Visible = False
End If
End Sub

Private Sub cmdSwitchUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If frmMain.cmdLogOut.Visible = True Then
        frmMain.Show
        Unload Me
        frmMain.cmdLogOut.BackColor = &H82D1B0
        frmMain.Enabled = True
    End If
    
    If frmMain.cmdManageAccounts.Caption = "EXIT" Then
        frmMain.Show
        Unload Me
        frmMain.cmdManageAccounts.BackColor = &H82D1B0
        frmMain.Enabled = True
    End If
End If
End Sub
