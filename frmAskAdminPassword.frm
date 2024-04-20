VERSION 5.00
Begin VB.Form frmAskAdminPassword 
   BackColor       =   &H003E472C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   8700
   ClientTop       =   4200
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton Exit 
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
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdAskOk 
         BackColor       =   &H00DBF2E9&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtAskAdminPass 
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
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "•"
         TabIndex        =   1
         Top             =   720
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   120
         Picture         =   "frmAskAdminPassword.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   375
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
         Left            =   2910
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER PASSWORD:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frmAskAdminPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public whatstatus As String
Public whatbutton As String

Private Sub cmdAskOk_Click() 'kailangan

'depende sa kung anong button ang klinick
If whatbutton = "addaccount" Then               'eto para sa add account
    If txtAskAdminPass.Text = Empty Then
        lblmessage1.Caption = "Please Enter Password!"
        lblmessage1.Visible = True
        txtAskAdminPass.SetFocus
    Else
        frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass='" + txtAskAdminPass.Text + "'"
        frmLogin.accountsdata.Refresh
        
        If frmLogin.accountsdata.Recordset.EOF Then
            lblmessage1.Caption = "Incorrect Password!"
            lblmessage1.Visible = True
            txtAskAdminPass.SetFocus
        Else
            txtAskAdminPass.Text = ""
            frmAddEdit.frmAddAccount.Visible = True
            frmAddEdit.frmUpdateAccount.Visible = False
            frmAddEdit.frmMyAccount.Visible = False
            Unload Me
            frmAddEdit.Show
            frmAddEdit.cmbUserType1.SetFocus
            
            'dito para magauto increment yung account id
            frmAddEdit.accountsdata.RecordSource = "select * from tblaccounts"
            frmAddEdit.accountsdata.Refresh
            
            Do
            frmAddEdit.accid = frmAddEdit.accountsdata.Recordset.Fields("accid")
            frmAddEdit.accountsdata.Recordset.MoveNext
            Loop Until frmAddEdit.accountsdata.Recordset.EOF
            
            frmAddEdit.accid = frmAddEdit.accid + 1
            frmAddEdit.lblaccid1.Caption = frmAddEdit.accid
            'hanggang dito lang
        End If
    End If
End If



If whatbutton = "updateaccount" Then            'para sa update account
    If txtAskAdminPass.Text = Empty Then
        lblmessage1.Caption = "Please Enter Password!"
        lblmessage1.Visible = True
        txtAskAdminPass.SetFocus
    Else
        frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
        frmLogin.accountsdata.Refresh
        
        If frmLogin.accountsdata.Recordset.EOF Then
            lblmessage1.Caption = "Incorrect Password!"
            lblmessage1.Visible = True
            txtAskAdminPass.SetFocus
        Else
            txtAskAdminPass.Text = ""
            frmAddEdit.frmUpdateAccount.Visible = True
            frmAddEdit.frmMyAccount.Visible = False
            frmAddEdit.frmAddAccount.Visible = False
            
            'para makuha yung selected data sa dbgrid papunta sa update account form
            frmAddEdit.lblaccid_2.Caption = frmMain.accountsdata.Recordset.Fields("accid")
            frmAddEdit.lblusername_2.Caption = frmMain.accountsdata.Recordset.Fields("user")
            frmAddEdit.cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype")
            frmAddEdit.txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first")
            frmAddEdit.txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle")
            frmAddEdit.txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last")
            frmAddEdit.txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age")
            frmAddEdit.dtbday2 = frmMain.accountsdata.Recordset.Fields("bday")
            frmAddEdit.cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex")
            frmAddEdit.txtAddress2 = frmMain.accountsdata.Recordset.Fields("address")
            frmAddEdit.txtPhone2 = frmMain.accountsdata.Recordset.Fields("phone")
            frmAddEdit.txtEmail2 = frmMain.accountsdata.Recordset.Fields("email")
            whatstatus = frmMain.accountsdata.Recordset.Fields("status")
            
            frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "'"
            frmLogin.accountsdata.Refresh
            
            If whatstatus = "ACTIVE" Then
                frmAddEdit.cmdstatus.Caption = "DEACTIVATE"
                frmAddEdit.cmdstatus.BackColor = &HC0C0FF
            Else
                frmAddEdit.cmdstatus.Caption = "ACTIVATE"
                frmAddEdit.cmdstatus.BackColor = &HC0FFC0
            End If
            
            If frmMain.accountsdata.Recordset.Fields("user") = frmLogin.accountsdata.Recordset.Fields("user") Then
                frmAddEdit.cmbUserType2.Enabled = False
                frmAddEdit.lblmessage2_7.Caption = "(Unable To Change User Type)"
                frmAddEdit.cmdstatus.Visible = False
                frmAddEdit.Show
                Unload Me
                frmAddEdit.txtFirst2.SetFocus
                frmAddEdit.lblmessage2_7.Visible = True
            Else
                frmAddEdit.Show
                Unload Me
                frmAddEdit.cmbUserType2.SetFocus
            End If
        End If
    End If
End If



If whatbutton = "myaccount" Then                'para sa my account
    If txtAskAdminPass.Text = Empty Then
        lblmessage1.Caption = "Please Enter Password!"
        lblmessage1.Visible = True
        txtAskAdminPass.SetFocus
    Else
        If Not frmLogin.txtUsername1.Text = Empty Then 'para sa admin accounts
            frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
            frmLogin.accountsdata.Refresh
            
            If frmLogin.accountsdata.Recordset.EOF Then
                lblmessage1.Caption = "Incorrect Password!"
                lblmessage1.Visible = True
                txtAskAdminPass.SetFocus
            Else
                'dito yung pagpumasok yung account ng admin
                frmAddEdit.lblid10.Caption = frmLogin.accountsdata.Recordset.Fields("accid")
                frmAddEdit.lblusername10.Caption = frmLogin.accountsdata.Recordset.Fields("user")
                frmAddEdit.lblusertype10.Caption = frmLogin.accountsdata.Recordset.Fields("usertype")
                frmAddEdit.lblfirst10.Caption = frmLogin.accountsdata.Recordset.Fields("first")
                frmAddEdit.lblmiddle10.Caption = frmLogin.accountsdata.Recordset.Fields("middle")
                frmAddEdit.lbllast10.Caption = frmLogin.accountsdata.Recordset.Fields("last")
                frmAddEdit.lblage10.Caption = frmLogin.accountsdata.Recordset.Fields("age")
                frmAddEdit.lblsex10.Caption = frmLogin.accountsdata.Recordset.Fields("sex")
                frmAddEdit.lblbday10.Caption = frmLogin.accountsdata.Recordset.Fields("bday")
                frmAddEdit.lblphone10.Caption = frmLogin.accountsdata.Recordset.Fields("phone")
                frmAddEdit.lblemail10.Caption = frmLogin.accountsdata.Recordset.Fields("email")
                frmAddEdit.lbladdress10.Caption = frmLogin.accountsdata.Recordset.Fields("address")
                
                
                
                frmAddEdit.frmMyAccount.Visible = True
                frmAddEdit.Show
                frmAddEdit.txtcurrentpass.SetFocus
                Unload Me
            End If
        Else
            If Not frmLogin.txtUsername2.Text = Empty Then 'para sa users accounts
                frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername2.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
                frmLogin.accountsdata.Refresh
                
                If frmLogin.accountsdata.Recordset.EOF Then
                    lblmessage1.Caption = "Incorrect Password!"
                    lblmessage1.Visible = True
                    txtAskAdminPass.SetFocus
                Else
                    'dito yung pagpumasok yung account ng user
                    frmAddEdit.lblid10.Caption = frmLogin.accountsdata.Recordset.Fields("accid")
                    frmAddEdit.lblusername10.Caption = frmLogin.accountsdata.Recordset.Fields("user")
                    frmAddEdit.lblusertype10.Caption = frmLogin.accountsdata.Recordset.Fields("usertype")
                    frmAddEdit.lblfirst10.Caption = frmLogin.accountsdata.Recordset.Fields("first")
                    frmAddEdit.lblmiddle10.Caption = frmLogin.accountsdata.Recordset.Fields("middle")
                    frmAddEdit.lbllast10.Caption = frmLogin.accountsdata.Recordset.Fields("last")
                    frmAddEdit.lblage10.Caption = frmLogin.accountsdata.Recordset.Fields("age")
                    frmAddEdit.lblsex10.Caption = frmLogin.accountsdata.Recordset.Fields("sex")
                    frmAddEdit.lblbday10.Caption = frmLogin.accountsdata.Recordset.Fields("bday")
                    frmAddEdit.lblphone10.Caption = frmLogin.accountsdata.Recordset.Fields("phone")
                    frmAddEdit.lblemail10.Caption = frmLogin.accountsdata.Recordset.Fields("email")
                    frmAddEdit.lbladdress10.Caption = frmLogin.accountsdata.Recordset.Fields("address")
                    
                    
                    
                    frmAddEdit.frmMyAccount.Visible = True
                    frmAddEdit.Show
                    frmAddEdit.txtcurrentpass.SetFocus
                    Unload Me
                End If
            End If
        End If
    End If
End If



If whatbutton = "deleteaccount" Then                        'para sa delete account button
    Dim reply As String
    
    If txtAskAdminPass.Text = Empty Then
        lblmessage1.Caption = "Please Enter Password!"
        lblmessage1.Visible = True
        txtAskAdminPass.SetFocus
    Else
        frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
        frmLogin.accountsdata.Refresh
        
        If frmLogin.accountsdata.Recordset.EOF Then
            lblmessage1.Caption = "Incorrect Password!"
            lblmessage1.Visible = True
            txtAskAdminPass.SetFocus
        Else
            frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "'"
            frmLogin.accountsdata.Refresh
            
            If frmMain.accountsdata.Recordset.Fields("accid") = frmLogin.accountsdata.Recordset("accid") Then
                frmMain.Show
                Unload Me
                frmMain.Enabled = True
                MsgBox "Account Is Currently In Use!", vbCritical, ""
                frmMain.cmdDeleteAdmin.BackColor = &H8080FF
            Else
                frmMain.Show
                Unload Me
                frmMain.Enabled = True
                reply = MsgBox("Delete The Account (" & frmMain.accountsdata.Recordset.Fields("accid") & " - " & frmMain.accountsdata.Recordset.Fields("last") & ") Permanently?", vbExclamation + vbYesNo, "")
                        
                If reply = vbYes Then
                    frmMain.accountsdata.Recordset.Delete
                    MsgBox "Account Deleted!", vbInformation, ""
                    
                    frmMain.accountsdata.Refresh
                End If
                
                frmMain.cmdDeleteAdmin.BackColor = &H8080FF
            End If
            
            frmMain.txtsearchaccount.Text = ""
        End If
    End If
End If


If whatbutton = "viewaccount" Then
    If txtAskAdminPass.Text = Empty Then
        lblmessage1.Caption = "Please Enter Password!"
        lblmessage1.Visible = True
        txtAskAdminPass.SetFocus
    Else
        frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
        frmLogin.accountsdata.Refresh
        
        If frmLogin.accountsdata.Recordset.EOF Then
            lblmessage1.Caption = "Incorrect Password!"
            lblmessage1.Visible = True
            txtAskAdminPass.SetFocus
        Else
            txtAskAdminPass.Text = ""
            frmView.frmViewAccount.Visible = True
            
            frmView.lblaccountid1.Caption = frmMain.accountsdata.Recordset.Fields("accid")
            frmView.lblusername1.Caption = frmMain.accountsdata.Recordset.Fields("user")
            frmView.lblfirst1.Caption = frmMain.accountsdata.Recordset.Fields("first")
            frmView.lblmiddle1.Caption = frmMain.accountsdata.Recordset.Fields("middle")
            frmView.lbllast1.Caption = frmMain.accountsdata.Recordset.Fields("last")
            frmView.lblage1.Caption = frmMain.accountsdata.Recordset.Fields("age")
            frmView.lblbday1.Caption = frmMain.accountsdata.Recordset.Fields("bday")
            frmView.lblsex1.Caption = frmMain.accountsdata.Recordset.Fields("sex")
            frmView.lblphone1.Caption = frmMain.accountsdata.Recordset.Fields("phone")
            frmView.lblemail1.Caption = frmMain.accountsdata.Recordset.Fields("email")
            
            If frmMain.accountsdata.Recordset.Fields("status") = "ACTIVE" Then
                frmView.lblstatus1.ForeColor = &H4000&
                frmView.lblstatus1.Caption = frmMain.accountsdata.Recordset.Fields("status")
            Else
                frmView.lblstatus1.ForeColor = &HFF&
                frmView.lblstatus1.Caption = frmMain.accountsdata.Recordset.Fields("status")
            End If
            
            If frmMain.accountsdata.Recordset.Fields("usertype") = "ADMIN" Then
                frmView.lblusertype1.Caption = frmMain.accountsdata.Recordset.Fields("usertype")
                frmView.imgusertype1.Picture = LoadPicture(App.Path & "/images/admin_logo_3.jpg")
            Else
                frmView.lblusertype1.Caption = frmMain.accountsdata.Recordset.Fields("usertype")
                frmView.imgusertype1.Picture = LoadPicture(App.Path & "/images/user_logo_3.jpg")
            End If
            
            frmView.Show
            Unload Me
        End If
    End If
End If



If whatbutton = "addproduct" Then
    If txtAskAdminPass.Text = Empty Then
        lblmessage1.Caption = "Please Enter Password!"
        lblmessage1.Visible = True
        txtAskAdminPass.SetFocus
    Else
        frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
        frmLogin.accountsdata.Refresh
        
        If frmLogin.accountsdata.Recordset.EOF Then
            lblmessage1.Caption = "Incorrect Password!"
            lblmessage1.Visible = True
            txtAskAdminPass.SetFocus
        Else
            txtAskAdminPass.Text = ""
            frmView.frmAddProduct.Visible = True
            frmView.Show
            Unload Me
            
            If frmView.productsdatainv.Recordset.RecordCount = 0 Then
                frmView.prodid = 1
                frmView.lblprodid1.Caption = frmView.prodid
            Else
                frmView.productsdatainv.RecordSource = "select * from tblproducts"
                frmView.productsdatainv.Refresh
                
                Do
                frmView.prodid = frmView.productsdatainv.Recordset.Fields("productid")
                frmView.productsdatainv.Recordset.MoveNext
                Loop Until frmView.productsdatainv.Recordset.EOF
                    
                frmView.prodid = frmView.prodid + 1
                frmView.lblprodid1.Caption = frmView.prodid
            End If

        End If
    End If
End If



If whatbutton = "updateproduct" Then
    If txtAskAdminPass.Text = Empty Then
            lblmessage1.Caption = "Please Enter Password!"
            lblmessage1.Visible = True
            txtAskAdminPass.SetFocus
        Else
            frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
            frmLogin.accountsdata.Refresh
                
            If frmLogin.accountsdata.Recordset.EOF Then
                lblmessage1.Caption = "Incorrect Password!"
                lblmessage1.Visible = True
                txtAskAdminPass.SetFocus
            Else
                frmView.Show
                frmView.frmUpdateProduct.Visible = True
                Unload Me
            
                frmView.lblprodid2.Caption = frmMain.productsdatainv.Recordset.Fields("productid")
                frmView.txtbarcode2.Text = frmMain.productsdatainv.Recordset.Fields("barcode")
                frmView.txtprodname2.Text = frmMain.productsdatainv.Recordset.Fields("productname")
                frmView.txtprice2.Text = frmMain.productsdatainv.Recordset.Fields("price")
                frmView.txtdescription2.Text = frmMain.productsdatainv.Recordset.Fields("description")
            End If
    End If
End If


If whatbutton = "deleteproduct" Then
    Dim reply2 As String
    
    If txtAskAdminPass.Text = Empty Then
        lblmessage1.Caption = "Please Enter Password!"
        lblmessage1.Visible = True
        txtAskAdminPass.SetFocus
    Else
        frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
        frmLogin.accountsdata.Refresh
            
        If frmLogin.accountsdata.Recordset.EOF Then
            lblmessage1.Caption = "Incorrect Password!"
            lblmessage1.Visible = True
            txtAskAdminPass.SetFocus
        Else
            Unload Me
            frmMain.Show
            frmMain.Enabled = True
            
            reply2 = MsgBox("Delete This Product (" & frmMain.productsdatainv.Recordset.Fields("productid") & " - " & frmMain.productsdatainv.Recordset.Fields("productname") & ") Permanently?", vbExclamation + vbYesNo, "")
            
            If reply2 = vbYes Then
                frmMain.productsdatainv.Recordset.Delete
                MsgBox "Product Deleted!", vbInformation, ""
                'frmMain.productsdatainv.Refresh
                
                frmMain.stocksdata.RecordSource = "select * from tblstocks where productid= '" + frmMain.idno + "'"
                frmMain.stocksdata.Refresh
                
                If Not frmMain.stocksdata.Recordset.EOF Then
                    frmMain.stocksdata.Recordset.Delete
                End If
                
                frmMain.productsdatainv.RecordSource = "select * from tblproducts"
                frmMain.productsdatainv.Refresh
            End If
            
            frmMain.txtsearchproduct_2.SetFocus
        End If
    End If
End If


If whatbutton = "addnewstock" Then
   If txtAskAdminPass.Text = Empty Then
        lblmessage1.Caption = "Please Enter Password!"
        lblmessage1.Visible = True
        txtAskAdminPass.SetFocus
    Else
        If Not frmLogin.txtUsername1.Text = Empty Then 'para sa admin accounts
            frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
            frmLogin.accountsdata.Refresh
            
            If frmLogin.accountsdata.Recordset.EOF Then
                lblmessage1.Caption = "Incorrect Password!"
                lblmessage1.Visible = True
                txtAskAdminPass.SetFocus
            Else
                'dito yung pagpumasok yung account ng admin
                frmView.Show
                frmView.frmAddNewStock.Visible = True
                frmView.cmbProducts.SetFocus
                Unload Me
            End If
        Else
            If Not frmLogin.txtUsername2.Text = Empty Then 'para sa users accounts
                frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername2.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
                frmLogin.accountsdata.Refresh
                
                If frmLogin.accountsdata.Recordset.EOF Then
                    lblmessage1.Caption = "Incorrect Password!"
                    lblmessage1.Visible = True
                    txtAskAdminPass.SetFocus
                Else
                    'dito yung pagpumasok yung account ng user
                    frmView.Show
                    frmView.frmAddNewStock.Visible = True
                    frmView.cmbProducts.SetFocus
                    Unload Me
                End If
            End If
        End If
    End If
End If



If whatbutton = "deletestock" Then
    Dim reply3 As String
    If txtAskAdminPass.Text = Empty Then
        lblmessage1.Caption = "Please Enter Password!"
        lblmessage1.Visible = True
        txtAskAdminPass.SetFocus
    Else
        If Not frmLogin.txtUsername1.Text = Empty Then 'para sa admin accounts
            frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
            frmLogin.accountsdata.Refresh
            
            If frmLogin.accountsdata.Recordset.EOF Then
                lblmessage1.Caption = "Incorrect Password!"
                lblmessage1.Visible = True
                txtAskAdminPass.SetFocus
            Else
                'dito yung pagpumasok yung account ng admin
                frmMain.Show
                frmMain.Enabled = True
                Unload Me
                
                reply3 = MsgBox("Delete This Stock (" & frmMain.stocksdata.Recordset.Fields("productid") & " - " & frmMain.stocksdata.Recordset.Fields("productname") & ") Permanently?", vbExclamation + vbYesNo, "")
                
                If reply3 = vbYes Then
                    MsgBox "Stock Deleted!", vbInformation, ""
                    frmMain.stocksdata.Recordset.Delete
                    frmMain.stocksdata.Refresh
                End If
                
                frmMain.txtsearchproduct.SetFocus
            End If
        Else
            If Not frmLogin.txtUsername2.Text = Empty Then 'para sa users accounts
                frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername2.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
                frmLogin.accountsdata.Refresh
                
                If frmLogin.accountsdata.Recordset.EOF Then
                    lblmessage1.Caption = "Incorrect Password!"
                    lblmessage1.Visible = True
                    txtAskAdminPass.SetFocus
                Else
                    'dito yung pagpumasok yung account ng user
                    frmMain.Show
                    frmMain.Enabled = True
                    Unload Me
                    
                    reply3 = MsgBox("Delete This Stock (" & frmMain.stocksdata.Recordset.Fields("productid") & " - " & frmMain.stocksdata.Recordset.Fields("productname") & ") Permanently?", vbExclamation + vbYesNo, "")
                    
                    If reply3 = vbYes Then
                        MsgBox "Stock Deleted!", vbInformation, ""
                        frmMain.stocksdata.Recordset.Delete
                        frmMain.stocksdata.Refresh
                    End If
                    
                    frmMain.txtsearchproduct.SetFocus
                End If
            End If
        End If
    End If
End If



If whatbutton = "updatestockqty" Then
    If txtAskAdminPass.Text = Empty Then
        lblmessage1.Caption = "Please Enter Password!"
        lblmessage1.Visible = True
        txtAskAdminPass.SetFocus
    Else
        If Not frmLogin.txtUsername1.Text = Empty Then 'para sa admin accounts
            frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
            frmLogin.accountsdata.Refresh
            
            If frmLogin.accountsdata.Recordset.EOF Then
                lblmessage1.Caption = "Incorrect Password!"
                lblmessage1.Visible = True
                txtAskAdminPass.SetFocus
            Else
                'dito yung pagpumasok yung account ng admin
                frmView.Show
                frmView.frmUpdateStock.Visible = True
                Unload Me
                
                frmView.lblprodid5.Caption = frmMain.stocksdata.Recordset.Fields("productid")
                frmView.lblbarcode5.Caption = frmMain.stocksdata.Recordset.Fields("barcode")
                frmView.lblprodname5.Caption = frmMain.stocksdata.Recordset.Fields("productname")
                frmView.lblprice5.Caption = frmMain.stocksdata.Recordset.Fields("price")
                frmView.txtdescription5.Text = frmMain.stocksdata.Recordset.Fields("description")
                frmView.txtstocks5.Text = frmMain.stocksdata.Recordset.Fields("stocks")
            End If
        Else
            If Not frmLogin.txtUsername2.Text = Empty Then 'para sa users accounts
                frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername2.Text + "' and pass= '" + txtAskAdminPass.Text + "'"
                frmLogin.accountsdata.Refresh
                
                If frmLogin.accountsdata.Recordset.EOF Then
                    lblmessage1.Caption = "Incorrect Password!"
                    lblmessage1.Visible = True
                    txtAskAdminPass.SetFocus
                Else
                    'dito yung pagpumasok yung account ng user
                    frmView.Show
                    frmView.frmUpdateStock.Visible = True
                    Unload Me
                    
                    frmView.lblprodid5.Caption = frmMain.stocksdata.Recordset.Fields("productid")
                    frmView.lblbarcode5.Caption = frmMain.stocksdata.Recordset.Fields("barcode")
                    frmView.lblprodname5.Caption = frmMain.stocksdata.Recordset.Fields("productname")
                    frmView.lblprice5.Caption = frmMain.stocksdata.Recordset.Fields("price")
                    frmView.txtdescription5.Text = frmMain.stocksdata.Recordset.Fields("description")
                    frmView.txtstocks5.Text = frmMain.stocksdata.Recordset.Fields("stocks")
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub cmdAskOk_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
frmMain.Show
Unload Me
frmMain.cmdMyAccount.BackColor = &H82D1B0
frmMain.cmdAddAdmin.BackColor = &H82D1B0
frmMain.cmdUpdateAdmin.BackColor = &H82D1B0
frmMain.cmdView.BackColor = &H82D1B0
frmMain.cmdDeleteAdmin.BackColor = &H8080FF
frmMain.Enabled = True
End If
End Sub

Private Sub Exit_Click() 'kailangan

frmMain.cmdAddAdmin.BackColor = &H82D1B0
frmMain.cmdUpdateAdmin.BackColor = &H82D1B0
frmMain.cmdView.BackColor = &H82D1B0
frmMain.cmdDeleteAdmin.BackColor = &H8080FF
frmMain.Enabled = True
frmMain.Show
Unload Me

If whatbutton = "myaccount" Then
    frmMain.cmdMyAccount.BackColor = &H82D1B0
    frmMain.Enabled = True
    Unload Me
End If

End Sub

Private Sub Exit_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
frmMain.Show
Unload Me
frmMain.cmdMyAccount.BackColor = &H82D1B0
frmMain.cmdAddAdmin.BackColor = &H82D1B0
frmMain.cmdUpdateAdmin.BackColor = &H82D1B0
frmMain.cmdView.BackColor = &H82D1B0
frmMain.cmdDeleteAdmin.BackColor = &H8080FF
frmMain.Enabled = True
End If
End Sub

Private Sub txtAskAdminPass_KeyPress(KeyAscii As Integer) 'kailangan
If KeyAscii = 13 Then cmdAskOk.SetFocus
If KeyAscii = 8 Then Exit Sub
If KeyAscii = 27 Then
frmMain.Show
Unload Me
frmMain.cmdMyAccount.BackColor = &H82D1B0
frmMain.cmdAddAdmin.BackColor = &H82D1B0
frmMain.cmdUpdateAdmin.BackColor = &H82D1B0
frmMain.cmdView.BackColor = &H82D1B0
frmMain.cmdDeleteAdmin.BackColor = &H8080FF
frmMain.Enabled = True
End If
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789._", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtAskAdminPass_Change() 'kailangan
lblmessage1.Visible = False
End Sub
