VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H004E8E6A&
   BorderStyle     =   0  'None
   Caption         =   "Admin Login"
   ClientHeight    =   6375
   ClientLeft      =   2790
   ClientTop       =   3090
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3766.56
   ScaleMode       =   0  'User
   ScaleWidth      =   11520.86
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H008080FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11955
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Data accountsdata 
      Caption         =   "ALL ACCS."
      Connect         =   "Access"
      DatabaseName    =   "database\database_sabana.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblaccounts"
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame frmadminlogin 
      BackColor       =   &H00EAF1F4&
      BorderStyle     =   0  'None
      Caption         =   "Admin Accts."
      Height          =   5655
      Left            =   3075
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdCancel1 
         BackColor       =   &H0082D1B0&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3345
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3960
         Width           =   1140
      End
      Begin VB.CommandButton cmdLogin1 
         BackColor       =   &H0082D1B0&
         Caption         =   "LOG IN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1665
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3960
         Width           =   1140
      End
      Begin VB.TextBox txtPassword1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "•"
         TabIndex        =   9
         Top             =   3120
         Width           =   3045
      End
      Begin VB.TextBox txtUsername1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   8
         Top             =   2280
         Width           =   3045
      End
      Begin VB.Image Image5 
         Height          =   375
         Left            =   120
         Picture         =   "frmLogin.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   2520
         Picture         =   "frmLogin.frx":1F42F
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADMIN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2535
         TabIndex        =   15
         Top             =   75
         Width           =   1080
      End
      Begin VB.Label lblmessage1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   3090
         TabIndex        =   13
         Top             =   3600
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   2880
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   1560
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H0082D1B0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0082D1B0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.Frame frmUserType 
      BackColor       =   &H00EAF1F4&
      BorderStyle     =   0  'None
      Caption         =   "Admin Accts."
      Height          =   5655
      Left            =   3075
      TabIndex        =   21
      Top             =   480
      Width           =   6135
      Begin VB.CommandButton cmdAdmin 
         BackColor       =   &H0082D1B0&
         Caption         =   "ADMIN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdUser 
         BackColor       =   &H0082D1B0&
         Caption         =   "USERS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Image Image7 
         Height          =   375
         Left            =   120
         Picture         =   "frmLogin.frx":2A137
         Stretch         =   -1  'True
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT USER TYPE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1890
         TabIndex        =   22
         Top             =   90
         Width           =   2325
      End
      Begin VB.Image Image4 
         Height          =   1455
         Left            =   2280
         Picture         =   "frmLogin.frx":49566
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   1455
         Left            =   2280
         Picture         =   "frmLogin.frx":5426E
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0082D1B0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0082D1B0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.Frame frmuserlogin 
      BackColor       =   &H00EAF1F4&
      BorderStyle     =   0  'None
      Caption         =   "Admin Accts."
      Height          =   5655
      Left            =   3068
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtUsername2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         TabIndex        =   4
         Top             =   2280
         Width           =   3045
      End
      Begin VB.TextBox txtPassword2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "•"
         TabIndex        =   5
         Top             =   3120
         Width           =   3045
      End
      Begin VB.CommandButton cmdLogin2 
         BackColor       =   &H0082D1B0&
         Caption         =   "LOG IN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3960
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel2 
         BackColor       =   &H0082D1B0&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3345
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3960
         Width           =   1140
      End
      Begin VB.Image Image6 
         Height          =   375
         Left            =   120
         Picture         =   "frmLogin.frx":6118D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   20
         Top             =   75
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "USERNAME:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1560
         TabIndex        =   18
         Top             =   2880
         Width           =   1125
      End
      Begin VB.Label lblmessage2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   3090
         TabIndex        =   17
         Top             =   3600
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image Image2 
         Height          =   1080
         Left            =   2520
         Picture         =   "frmLogin.frx":805BC
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1080
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0082D1B0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0082D1B0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   0
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H003E472C&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   -105
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdadmin_Click()
    typeuser = "ADMIN"
    frmadminlogin.Visible = True
    frmuserlogin.Visible = False
    frmUserType.Visible = False
    txtUsername1.SetFocus
End Sub

Private Sub cmdCancel1_Click() 'admin
    txtUsername1.Text = ""
    txtPassword1.Text = ""
    frmUserType.Visible = True
    frmadminlogin.Visible = False
End Sub

Private Sub cmdCancel1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmadminlogin.Visible = False
    frmuserlogin.Visible = False
    frmUserType.Visible = True
    txtUsername1.Text = ""
    txtPassword1.Text = ""
    cmdAdmin.SetFocus
End If
End Sub

Private Sub cmdCancel2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmadminlogin.Visible = False
    frmuserlogin.Visible = False
    frmUserType.Visible = True
    txtUsername2.Text = ""
    txtPassword2.Text = ""
    cmdAdmin.SetFocus
End If
End Sub

Private Sub cmdExit_Click()
Dim reply As String

reply = MsgBox("Exit Program?", vbQuestion + vbYesNo, "")

If reply = vbYes Then
    End
End If
End Sub

Private Sub cmdLogin1_Click() 'admin
'scan database of admin.
On Error Resume Next

accountsdata.RecordSource = "select * from tblaccounts where user= '" + txtUsername1.Text + "' and pass= '" + txtPassword1.Text + "'"
accountsdata.Refresh

If accountsdata.Recordset.EOF Then
    If txtUsername1.Text = Empty Or txtPassword1.Text = Empty Then
        lblmessage1.Caption = "Please Enter Username And Password!"
        lblmessage1.Visible = True
        txtUsername1.SetFocus
    Else
        lblmessage1.Caption = "Account Doesn't Exist!"
        lblmessage1.Visible = True
        txtUsername1.SetFocus
    End If
Else
    If accountsdata.Recordset.Fields("usertype") = "ADMIN" Then
        If accountsdata.Recordset.Fields("status") = "ACTIVE" Then
            MsgBox "Welcome " & accountsdata.Recordset.Fields("last") & "!", vbInformation, ""
            frmMain.Show
            frmLogin.Hide
            frmMain.menubutton = "inventoryadmin"
            frmMain.lblfirst.Caption = accountsdata.Recordset.Fields("first")
            frmMain.lbllast.Caption = accountsdata.Recordset.Fields("last")
            frmMain.lblusertype.Caption = accountsdata.Recordset.Fields("usertype")
            frmMain.imgUserType.Picture = LoadPicture(App.Path & "/images/admin_logo_2.jpg")
            frmMain.txtsearchproduct_3.SetFocus
        Else
            lblmessage1.Caption = "Account Doesn't Exist!"
            lblmessage1.Visible = True
            txtUsername1.SetFocus
        End If
    Else
        lblmessage1.Caption = "Account Is Not Admin!"
        lblmessage1.Visible = True
        txtUsername1.SetFocus
    End If
End If
End Sub

Private Sub cmdLogin1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmadminlogin.Visible = False
    frmuserlogin.Visible = False
    frmUserType.Visible = True
    txtUsername1.Text = ""
    txtPassword1.Text = ""
    cmdAdmin.SetFocus
End If
End Sub

Private Sub cmdLogin2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmadminlogin.Visible = False
    frmuserlogin.Visible = False
    frmUserType.Visible = True
    txtUsername2.Text = ""
    txtPassword2.Text = ""
    cmdAdmin.SetFocus
End If
End Sub

Private Sub cmdUser_Click()
    typeuser = "USER"
    frmuserlogin.Visible = True
    frmadminlogin.Visible = False
    frmUserType.Visible = False
    txtUsername2.SetFocus
End Sub



Private Sub txtPassword1_Change() 'admin
    lblmessage1.Visible = False
End Sub

Private Sub txtPassword1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdLogin1.SetFocus
If KeyAscii = 27 Then
    frmadminlogin.Visible = False
    frmuserlogin.Visible = False
    frmUserType.Visible = True
    txtUsername1.Text = ""
    txtPassword1.Text = ""
    cmdAdmin.SetFocus
End If
End Sub

Private Sub txtPassword2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdLogin2.SetFocus
If KeyAscii = 27 Then
    frmadminlogin.Visible = False
    frmuserlogin.Visible = False
    frmUserType.Visible = True
    txtUsername2.Text = ""
    txtPassword2.Text = ""
    cmdAdmin.SetFocus
End If
End Sub

Private Sub txtUsername1_Change() 'admin
    lblmessage1.Visible = False
End Sub

Private Sub cmdCancel2_Click() 'users
    txtUsername2.Text = ""
    txtPassword2.Text = ""
    frmUserType.Visible = True
    frmuserlogin.Visible = False
End Sub

Private Sub cmdLogin2_Click() 'users
'scan database of users.
On Error Resume Next

accountsdata.RecordSource = "select * from tblaccounts where user= '" + txtUsername2.Text + "' and pass= '" + txtPassword2.Text + "'"
accountsdata.Refresh
        
If accountsdata.Recordset.EOF Then
    If txtUsername2.Text = Empty Or txtPassword2.Text = Empty Then
        lblmessage2.Caption = "Please Enter Username And Password!"
        lblmessage2.Visible = True
        txtUsername2.SetFocus
    Else
        lblmessage2.Caption = "Account Doesn't Exist!"
        lblmessage2.Visible = True
        txtUsername2.SetFocus
    End If
Else
    If accountsdata.Recordset.Fields("usertype") = "USER" Then
        If accountsdata.Recordset.Fields("status") = "ACTIVE" Then
            MsgBox "Welcome " & accountsdata.Recordset.Fields("last") & "!", vbInformation, ""
            frmMain.Show
            frmLogin.Hide
            frmMain.lblfirst.Caption = accountsdata.Recordset.Fields("first")
            frmMain.lbllast.Caption = accountsdata.Recordset.Fields("last")
            frmMain.cmdLogOut.Visible = False
            frmMain.menubutton = "inventoryuser"
            frmMain.cmdManageAccounts.Caption = "EXIT"
            frmMain.lblusertype.Caption = accountsdata.Recordset.Fields("usertype")
            frmMain.imgUserType.Picture = LoadPicture(App.Path & "/images/user_logo_2.jpg")
            frmMain.txtsearchproduct_3.SetFocus
        Else
            lblmessage2.Caption = "Account Doesn't Exist!"
            lblmessage2.Visible = True
            txtUsername2.SetFocus
        End If
    Else
        lblmessage2.Caption = "Account Is Not User!"
        lblmessage2.Visible = True
        txtUsername2.SetFocus
    End If
End If
End Sub


Private Sub txtPassword2_Change() 'users
    lblmessage2.Visible = False
End Sub

Private Sub txtUsername1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPassword1.SetFocus
If KeyAscii = 27 Then
    frmadminlogin.Visible = False
    frmuserlogin.Visible = False
    frmUserType.Visible = True
    txtUsername1.Text = ""
    txtPassword1.Text = ""
    cmdAdmin.SetFocus
End If
End Sub

Private Sub txtUserName2_Change() 'users
    lblmessage2.Visible = False
End Sub


Private Sub txtUsername2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPassword2.SetFocus
If KeyAscii = 27 Then
    frmadminlogin.Visible = False
    frmuserlogin.Visible = False
    frmUserType.Visible = True
    txtUsername2.Text = ""
    txtPassword2.Text = ""
    cmdAdmin.SetFocus
End If
End Sub
