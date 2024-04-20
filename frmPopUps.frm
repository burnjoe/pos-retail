VERSION 5.00
Begin VB.Form frmPopUps 
   BackColor       =   &H003E472C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmEditQty 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.ComboBox cmbEditType 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPopUps.frx":0000
         Left            =   4320
         List            =   "frmPopUps.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtEditQty 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   600
         MaxLength       =   6
         TabIndex        =   5
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label lbledittype 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADD QUANTITY:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRESS [ENTER]  -  PRESS [ESC] TO CANCEL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   6
         Top             =   2520
         Width           =   3225
      End
   End
   Begin VB.Frame frmAddQty 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox txtaddqty 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   600
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRESS [ENTER]  -  PRESS [ESC] TO CANCEL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   3
         Top             =   2520
         Width           =   3225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame frmSearchCart 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox txtSearchBarcode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1200
         Width           =   5535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER BARCODE:"
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
         Left            =   600
         TabIndex        =   12
         Top             =   600
         Width           =   2580
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRESS [ENTER]  -  PRESS [ESC] TO CANCEL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   11
         Top             =   2160
         Width           =   3225
      End
   End
   Begin VB.Frame frmPay 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox txtPay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   600
         MaxLength       =   14
         TabIndex        =   14
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRESS [ENTER]  -  PRESS [ESC] TO CANCEL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         TabIndex        =   16
         Top             =   2520
         Width           =   3225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CASH AMOUNT:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   15
         Top             =   360
         Width           =   2640
      End
   End
End
Attribute VB_Name = "frmPopUps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public subtotal, total, qtytotal, pay As Currency
Public date1, time1 As String
Public qty, ornumber As Integer


Private Sub cmbEditType_Click()
    If cmbEditType = "ADD" Then
        lbledittype.Caption = "ADD QUANTITY:"
        txtEditQty.Text = ""
    Else
        lbledittype.Caption = "DEDUCT QUANTITY:"
        txtEditQty.Text = frmMain.transactdata.Recordset.Fields("quantity")
        txtEditQty.SelStart = 0
        txtEditQty.SelLength = Len(txtEditQty)
    End If
txtEditQty.SetFocus
End Sub

Private Sub txtaddqty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(txtaddqty.Text) > Val(frmMain.stocksdata.Recordset.Fields("stocks")) Then
        MsgBox "Insufficient Stocks!", vbCritical, ""
        txtaddqty.Text = ""
        txtaddqty.SetFocus
    Else
        If txtaddqty.Text = Empty Then
            MsgBox "Enter Quantity!", vbCritical, ""
            txtaddqty.Text = ""
            txtaddqty.SetFocus
        Else
            If Val(txtaddqty.Text) = 0 Then
                MsgBox "Enter Quantity!", vbCritical, ""
                txtaddqty.Text = ""
                txtaddqty.SetFocus
            Else
                frmMain.transactdata.RecordSource = "select * from tbltransact where barcode= '" + frmMain.stocksdata.Recordset.Fields("barcode") + "'"
                frmMain.transactdata.Refresh
                
                If frmMain.transactdata.Recordset.EOF Then
                    frmMain.transactdata.Recordset.AddNew
                    frmMain.transactdata.Recordset.Fields("productid") = frmMain.stocksdata.Recordset.Fields("productid")
                    frmMain.transactdata.Recordset.Fields("barcode") = frmMain.stocksdata.Recordset.Fields("barcode")
                    frmMain.transactdata.Recordset.Fields("productname") = frmMain.stocksdata.Recordset.Fields("productname")
                    frmMain.transactdata.Recordset.Fields("price") = FormatNumber(frmMain.stocksdata.Recordset.Fields("price")) * Val(txtaddqty.Text)
                    frmMain.transactdata.Recordset.Fields("price") = FormatNumber(frmMain.transactdata.Recordset.Fields("price"))
                    frmMain.transactdata.Recordset.Fields("quantity") = Val(txtaddqty.Text)
                    subtotal = FormatNumber(frmMain.transactdata.Recordset.Fields("price"))
                    frmMain.transactdata.Recordset.Fields("quantityprice") = Val(frmMain.transactdata.Recordset.Fields("quantity")) & "  X  " & FormatNumber(frmMain.stocksdata.Recordset.Fields("price"))
                    frmMain.transactdata.Recordset.Update
                    
                    total = Val(total) + subtotal
                    frmMain.lbltotal.Caption = FormatNumber(total)
                    
                    If Len(frmMain.lbltotal) > 12 Then
                        frmMain.lbltotal.FontSize = 22
                        If Len(frmMain.lbltotal) > 13 Then
                            frmMain.lbltotal.FontSize = 18
                        Else
                            frmMain.lbltotal.FontSize = 22
                        End If
                    Else
                        frmMain.lbltotal.FontSize = 24
                    End If
                    
                    frmMain.stocksdata.Recordset.Edit
                    frmMain.stocksdata.Recordset.Fields("stocks") = Val(frmMain.stocksdata.Recordset.Fields("stocks")) - Val(txtaddqty.Text)
                    frmMain.stocksdata.Recordset.Update
                    
                    frmMain.transactdata.RecordSource = "select * from tbltransact"
                    frmMain.transactdata.Refresh
                    
                    frmMain.Show
                    frmMain.Enabled = True
                    Unload Me
                    frmMain.txtsearchproduct_3.Text = ""
                    frmMain.txtsearchproduct_3.SetFocus
                    
                    'SIDEMENU
                    frmMain.frmSideMenu.Enabled = False
                    frmMain.cmdMyAccount.BackColor = &HC0C0C0
                    frmMain.cmdTransaction.BackColor = &HC0C0C0
                    frmMain.cmdSales.BackColor = &HC0C0C0
                    frmMain.cmdInventory.BackColor = &HC0C0C0
                    frmMain.cmdManageAccounts.BackColor = &HC0C0C0
                    frmMain.cmdLogOut.BackColor = &HC0C0C0
                Else
                
                
                    frmMain.transactdata.Recordset.Edit
                    frmMain.transactdata.Recordset.Fields("quantity") = Val(frmMain.transactdata.Recordset.Fields("quantity")) + Val(txtaddqty.Text)
                    frmMain.transactdata.Recordset.Update
                    
                    total = Val(total) + (FormatNumber(frmMain.stocksdata.Recordset.Fields("price")) * Val(txtaddqty.Text))
                    frmMain.lbltotal.Caption = FormatNumber(total)
                    
                    If Len(frmMain.lbltotal) > 12 Then
                        frmMain.lbltotal.FontSize = 22
                        If Len(frmMain.lbltotal) > 13 Then
                            frmMain.lbltotal.FontSize = 18
                        Else
                            frmMain.lbltotal.FontSize = 22
                        End If
                    Else
                        frmMain.lbltotal.FontSize = 24
                    End If
                    
                    frmMain.transactdata.Recordset.Edit
                    frmMain.transactdata.Recordset.Fields("price") = (FormatNumber(frmMain.stocksdata.Recordset.Fields("price")) * Val(txtaddqty.Text)) + FormatNumber(frmMain.transactdata.Recordset.Fields("price"))
                    frmMain.transactdata.Recordset.Fields("price") = FormatNumber(frmMain.transactdata.Recordset.Fields("price"))
                    frmMain.transactdata.Recordset.Update
                    
                    frmMain.transactdata.Recordset.Edit
                    frmMain.transactdata.Recordset.Fields("quantityprice") = Val(frmMain.transactdata.Recordset.Fields("quantity")) & "  X  " & FormatNumber(frmMain.stocksdata.Recordset.Fields("price"))
                    frmMain.transactdata.Recordset.Update
                    
                    frmMain.stocksdata.Recordset.Edit
                    frmMain.stocksdata.Recordset.Fields("stocks") = Val(frmMain.stocksdata.Recordset.Fields("stocks")) - Val(txtaddqty.Text)
                    frmMain.stocksdata.Recordset.Update
                    
                    frmMain.transactdata.RecordSource = "select * from tbltransact"
                    frmMain.transactdata.Refresh
                    frmMain.Show
                    frmMain.Enabled = True
                    Unload Me
                    frmMain.txtsearchproduct_3.Text = ""
                    frmMain.txtsearchproduct_3.SetFocus
                    
                    'SIDE MENU
                    frmMain.frmSideMenu.Enabled = False
                    frmMain.cmdMyAccount.BackColor = &HC0C0C0
                    frmMain.cmdTransaction.BackColor = &HC0C0C0
                    frmMain.cmdSales.BackColor = &HC0C0C0
                    frmMain.cmdInventory.BackColor = &HC0C0C0
                    frmMain.cmdManageAccounts.BackColor = &HC0C0C0
                    frmMain.cmdLogOut.BackColor = &HC0C0C0
                End If
            End If
        End If
    End If
End If


If KeyAscii = 27 Then
    frmMain.Show
    frmMain.txtsearchproduct_3.Text = ""
    frmMain.stocksdata.RecordSource = "select * from tblstocks"
    frmMain.stocksdata.Refresh
    frmMain.transactdata.RecordSource = "select * from tbltransact"
    frmMain.transactdata.Refresh
    frmMain.Enabled = True
    Unload Me
    frmMain.txtsearchproduct_3.SetFocus
End If

If KeyAscii = 8 Then Exit Sub
If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtEditQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If cmbEditType.Text = "ADD" Then
        frmMain.stocksdata.RecordSource = "select * from tblstocks where barcode= '" + frmMain.transactdata.Recordset.Fields("barcode") + "'"
        frmMain.stocksdata.Refresh
        
        If Not frmMain.stocksdata.Recordset.EOF Then
            If Val(txtEditQty.Text) > Val(frmMain.stocksdata.Recordset.Fields("stocks")) Then
                MsgBox "Insufficient Stocks!", vbCritical, ""
                txtEditQty.Text = ""
                txtEditQty.SetFocus
            Else
                If txtEditQty.Text = Empty Then
                    MsgBox "Enter Quantity To Add!", vbCritical, ""
                    txtaddqty.Text = ""
                    txtEditQty.SetFocus
                Else
                    If txtEditQty.Text = 0 Then
                        MsgBox "Enter Quantity To Add!", vbCritical, ""
                        txtEditQty.Text = ""
                        txtEditQty.SetFocus
                    Else
                        
                        frmMain.transactdata.Recordset.Edit
                        frmMain.transactdata.Recordset.Fields("quantity") = frmMain.transactdata.Recordset.Fields("quantity") + Val(txtEditQty.Text)
                        frmMain.transactdata.Recordset.Update
                        
                        total = Val(total) + (Val(frmMain.stocksdata.Recordset.Fields("price")) * Val(txtEditQty.Text))
                        frmMain.lbltotal.Caption = FormatNumber(total)
                        
                        If Len(frmMain.lbltotal) > 12 Then
                            frmMain.lbltotal.FontSize = 22
                            If Len(frmMain.lbltotal) > 13 Then
                                frmMain.lbltotal.FontSize = 18
                            Else
                                frmMain.lbltotal.FontSize = 22
                            End If
                        Else
                            frmMain.lbltotal.FontSize = 24
                        End If
                        
                        frmMain.transactdata.Recordset.Edit
                        frmMain.transactdata.Recordset.Fields("price") = (FormatNumber(frmMain.stocksdata.Recordset.Fields("price")) * Val(txtEditQty.Text)) + FormatNumber(frmMain.transactdata.Recordset.Fields("price"))
                        frmMain.transactdata.Recordset.Fields("price") = Val(frmMain.transactdata.Recordset.Fields("price"))
                        frmMain.transactdata.Recordset.Update
                        
                        frmMain.transactdata.Recordset.Edit
                        frmMain.transactdata.Recordset.Fields("quantityprice") = frmMain.transactdata.Recordset.Fields("quantity") & "  X  " & FormatNumber(frmMain.stocksdata.Recordset.Fields("price"))
                        frmMain.transactdata.Recordset.Update
                        
                        frmMain.stocksdata.Recordset.Edit
                        frmMain.stocksdata.Recordset.Fields("stocks") = frmMain.stocksdata.Recordset.Fields("stocks") - Val(txtEditQty.Text)
                        frmMain.stocksdata.Recordset.Update
                        
                        frmMain.transactdata.RecordSource = "select * from tbltransact"
                        frmMain.transactdata.Refresh
                        frmMain.stocksdata.RecordSource = "select * from tblstocks"
                        frmMain.stocksdata.Refresh
                        frmMain.Show
                        frmMain.Enabled = True
                        Unload Me
                        frmMain.txtsearchproduct_3.Text = ""
                        frmMain.txtsearchproduct_3.SetFocus
                    End If
                End If
            End If
        End If
    Else
        If Val(txtEditQty.Text) = frmMain.transactdata.Recordset.Fields("quantity") Then
            total = Val(total) - FormatNumber(frmMain.transactdata.Recordset.Fields("price"))
            frmMain.lbltotal.Caption = FormatNumber(total)
            
            If Len(frmMain.lbltotal) > 12 Then
                frmMain.lbltotal.FontSize = 22
                If Len(frmMain.lbltotal) > 13 Then
                    frmMain.lbltotal.FontSize = 18
                Else
                    frmMain.lbltotal.FontSize = 22
                End If
            Else
                frmMain.lbltotal.FontSize = 24
            End If
            
            frmMain.stocksdata.RecordSource = "select * from tblstocks where barcode= '" + frmMain.transactdata.Recordset.Fields("barcode") + "'"
            frmMain.stocksdata.Refresh
            
            If Not frmMain.stocksdata.Recordset.EOF Then
                frmMain.stocksdata.Recordset.Edit
                frmMain.stocksdata.Recordset.Fields("stocks") = Val(frmMain.stocksdata.Recordset.Fields("stocks")) + Val(frmMain.transactdata.Recordset.Fields("quantity"))
                frmMain.stocksdata.Recordset.Update
                
                frmMain.stocksdata.RecordSource = "select * from tblstocks"
                frmMain.stocksdata.Refresh
            End If
            
            frmMain.transactdata.Recordset.Delete
            frmMain.transactdata.Refresh
            
            frmMain.transactdata.RecordSource = "select * from tbltransact"
            frmMain.transactdata.Refresh
            frmMain.stocksdata.RecordSource = "select * from tblstocks"
            frmMain.stocksdata.Refresh
            frmMain.Show
            frmMain.Enabled = True
            Unload Me
            frmMain.txtsearchproduct_3.Text = ""
            frmMain.txtsearchproduct_3.SetFocus
        Else
            frmMain.stocksdata.RecordSource = "select * from tblstocks where barcode= '" + frmMain.transactdata.Recordset.Fields("barcode") + "'"
            frmMain.stocksdata.Refresh
            
            If Not frmMain.stocksdata.Recordset.EOF Then
                If Val(txtEditQty.Text) > frmMain.transactdata.Recordset.Fields("quantity") Then
                    MsgBox "Quantity Entered Is Greater Than Quantity In Cart!", vbCritical, ""
                    txtEditQty.Text = frmMain.transactdata.Recordset.Fields("quantity")
                    txtEditQty.SetFocus
                    txtEditQty.SelStart = 0
                    txtEditQty.SelLength = Len(txtEditQty)
                Else
                    If txtEditQty.Text = Empty Then
                        MsgBox "Enter Quantity To Deduct!", vbCritical, ""
                        txtEditQty.Text = frmMain.transactdata.Recordset.Fields("quantity")
                        txtEditQty.SetFocus
                        txtEditQty.SelStart = 0
                        txtEditQty.SelLength = Len(txtEditQty)
                    Else
                        If txtEditQty.Text = 0 Then
                            MsgBox "Enter Quantity To Deduct!", vbCritical, ""
                            txtEditQty.Text = frmMain.transactdata.Recordset.Fields("quantity")
                            txtEditQty.SetFocus
                            txtEditQty.SelStart = 0
                            txtEditQty.SelLength = Len(txtEditQty)
                        Else
                            qtytotal = frmMain.transactdata.Recordset.Fields("quantity")
                            
                            frmMain.transactdata.Recordset.Edit
                            frmMain.transactdata.Recordset.Fields("quantity") = frmMain.transactdata.Recordset.Fields("quantity") - Val(txtEditQty.Text)
                            frmMain.transactdata.Recordset.Update
                            
                            'MALI
                            frmMain.transactdata.Recordset.Edit
                            frmMain.transactdata.Recordset.Fields("price") = FormatNumber(frmMain.stocksdata.Recordset.Fields("price")) * frmMain.transactdata.Recordset.Fields("quantity")
                            frmMain.transactdata.Recordset.Fields("price") = FormatNumber(frmMain.transactdata.Recordset.Fields("price"))
                            frmMain.transactdata.Recordset.Update
                            
                            total = FormatNumber(frmMain.transactdata.Recordset.Fields("price")) + (Val(total) - FormatNumber((frmMain.stocksdata.Recordset.Fields("price")) * qtytotal))
                            frmMain.lbltotal.Caption = FormatNumber(total)
                            'HANGGANG DITO
                            
                            If Len(frmMain.lbltotal) > 12 Then
                                frmMain.lbltotal.FontSize = 22
                                If Len(frmMain.lbltotal) > 13 Then
                                    frmMain.lbltotal.FontSize = 18
                                Else
                                    frmMain.lbltotal.FontSize = 22
                                End If
                            Else
                                frmMain.lbltotal.FontSize = 24
                            End If
                            
                            frmMain.transactdata.Recordset.Edit
                            frmMain.transactdata.Recordset.Fields("quantityprice") = frmMain.transactdata.Recordset.Fields("quantity") & "  X  " & FormatNumber(frmMain.stocksdata.Recordset.Fields("price"))
                            frmMain.transactdata.Recordset.Update
                            
                            frmMain.stocksdata.Recordset.Edit
                            frmMain.stocksdata.Recordset.Fields("stocks") = frmMain.stocksdata.Recordset.Fields("stocks") + Val(txtEditQty.Text)
                            frmMain.stocksdata.Recordset.Update
                            
                            frmMain.transactdata.RecordSource = "select * from tbltransact"
                            frmMain.transactdata.Refresh
                            frmMain.stocksdata.RecordSource = "select * from tblstocks"
                            frmMain.stocksdata.Refresh
                            frmMain.Show
                            frmMain.Enabled = True
                            Unload Me
                            frmMain.txtsearchproduct_3.Text = ""
                            frmMain.txtsearchproduct_3.SetFocus
                        End If
                    End If
                End If
            End If
        End If
    End If
    frmMain.cmdSearchCart.Caption = "SEARCH ITEM FROM CART"
    frmMain.cmdSearchCart.BackColor = &H82D1B0
End If

If KeyAscii = 27 Then
    frmMain.Show
    frmMain.txtsearchproduct_3.Text = ""
    frmMain.stocksdata.RecordSource = "select * from tblstocks"
    frmMain.stocksdata.Refresh
    frmMain.transactdata.RecordSource = "select * from tbltransact"
    frmMain.transactdata.Refresh
    frmMain.Enabled = True
    frmMain.cmdSearchCart.Caption = "SEARCH ITEM FROM CART"
    frmMain.cmdSearchCart.BackColor = &H82D1B0
    Unload Me
    frmMain.txtsearchproduct_3.SetFocus
End If

If KeyAscii = 8 Then Exit Sub
If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub


Private Sub txtPay_KeyPress(KeyAscii As Integer)
Dim answer2 As String

If KeyAscii = 13 Then
    If Val(txtPay.Text) < FormatNumber(frmMain.lbltotal.Caption) Then
        MsgBox "Insufficient Amount!", vbCritical, ""
        txtPay.Text = ""
        txtPay.SetFocus
    Else
        answer2 = MsgBox("Pay With This Amount? (" & FormatNumber(txtPay.Text) & ")", vbQuestion + vbYesNo, "")
        
        If answer2 = vbYes Then
            qty = 0
            ornumber = 0
            'overflow pay
            pay = Val(txtPay.Text)
            transactprint.Sections("Section5").Controls("lblcash").Caption = Format(Val(txtPay.Text), "0.00")
            frmMain.lblchange.Caption = Val(txtPay.Text) - Val(total)
            frmMain.lblchange.Caption = FormatNumber(frmMain.lblchange.Caption)
            
            frmMain.Show
            Unload Me
            
            'COMPUTE FOR ITEM PURCHASED
            Do
            qty = qty + Val(frmMain.transactdata.Recordset.Fields("quantity"))
            frmMain.transactdata.Recordset.MoveNext
            Loop Until frmMain.transactdata.Recordset.EOF
            
            date1 = Format(Date, "mmmm dd, yyyy")
            time1 = Time
                  
                    
            'PARA AUTO INCREMENT YUNG OR NUMBER
            frmMain.salesdata.RecordSource = "select * from tblsales"
            frmMain.salesdata.Refresh
            
            If frmMain.salesdata.Recordset.EOF Then
                ornumber = ornumber + 1
            Else
                Do
                ornumber = frmMain.salesdata.Recordset.Fields("orno")
                frmMain.salesdata.Recordset.MoveNext
                Loop Until frmMain.salesdata.Recordset.EOF
                    
                ornumber = ornumber + 1
            End If
            
            
            
            'PARA SA SALESDATA ACCESS
            With frmMain.salesdata.Recordset
                .AddNew
                .Fields("orno") = ornumber
                .Fields("cashier") = frmLogin.accountsdata.Recordset.Fields("accid") & "  " & frmLogin.accountsdata.Recordset.Fields("first")
                .Fields("date") = date1
                .Fields("time") = time1
                .Fields("itempurchased") = qty
                .Fields("cash") = FormatNumber(pay)
                .Fields("change") = FormatNumber(frmMain.lblchange.Caption)
                .Fields("total") = FormatNumber(frmMain.lbltotal.Caption)
                .Update
            End With
            frmMain.salesdata.Refresh
            
            
            
            'PARA SA PRODUCT SALES RECORD
            frmMain.transactdata.RecordSource = "select * from tbltransact"
            frmMain.transactdata.Refresh
            
            If Not frmMain.transactdata.Recordset.EOF Then
            Do
            With frmSalesRecord.productsalesdata.Recordset
                .AddNew
                .Fields("productid") = frmMain.transactdata.Recordset.Fields("productid")
                .Fields("orno") = ornumber
                .Fields("barcode") = frmMain.transactdata.Recordset.Fields("barcode")
                .Fields("productname") = frmMain.transactdata.Recordset.Fields("productname")
                .Fields("quantity") = frmMain.transactdata.Recordset.Fields("quantity")
                .Fields("price") = frmMain.transactdata.Recordset.Fields("price")
                .Fields("quantityprice") = frmMain.transactdata.Recordset.Fields("quantityprice")
                .Update
            End With
            frmMain.transactdata.Recordset.MoveNext
            Loop Until frmMain.transactdata.Recordset.EOF
            End If
            
            
            
            'PARA SA DATA REPORT TO
            transactprint.Sections("Section4").Controls("lblOR").Caption = ornumber
            transactprint.Sections("Section4").Controls("lblcashier").Caption = frmLogin.accountsdata.Recordset.Fields("accid") & "  " & frmLogin.accountsdata.Recordset.Fields("first")
            transactprint.Sections("Section4").Controls("lbldate").Caption = date1
            transactprint.Sections("Section4").Controls("lbltime").Caption = time1
            
            transactprint.Sections("Section5").Controls("lbltotal").Caption = FormatNumber(frmMain.lbltotal.Caption)
            transactprint.Sections("Section5").Controls("lblcash").Caption = FormatNumber(pay)
            transactprint.Sections("Section5").Controls("lblchange").Caption = FormatNumber(frmMain.lblchange.Caption)
            transactprint.Sections("Section5").Controls("lblquantity").Caption = qty
            
            Unload DataEnvironment1
            transactprint.Show
        End If
    End If
End If

If KeyAscii = 27 Then
    frmMain.Show
    frmMain.txtsearchproduct_3.Text = ""
    frmMain.stocksdata.RecordSource = "select * from tblstocks"
    frmMain.stocksdata.Refresh
    frmMain.transactdata.RecordSource = "select * from tbltransact"
    frmMain.transactdata.Refresh
    frmMain.Enabled = True
    frmMain.cmdSearchCart.Caption = "SEARCH ITEM FROM CART"
    frmMain.cmdSearchCart.BackColor = &H82D1B0
    Unload Me
    frmMain.txtsearchproduct_3.SetFocus
End If

If KeyAscii = 8 Then Exit Sub
If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtSearchBarcode_KeyPress(KeyAscii As Integer)

'dto ka huminto
If KeyAscii = 13 Then
    If frmMain.cmdSearchCart.Caption = "SEARCH ITEM FROM CART" Then
        frmMain.transactdata.RecordSource = "select * from tbltransact where barcode= '" + txtSearchBarcode.Text + "'"
        frmMain.transactdata.Refresh
        
        If frmMain.transactdata.Recordset.EOF Then
            MsgBox "Item Not Found!", vbCritical, ""
            txtSearchBarcode.Text = ""
            txtSearchBarcode.SetFocus
            frmMain.transactdata.RecordSource = "select * from tbltransact"
            frmMain.transactdata.Refresh
        Else
            frmMain.Show
            frmMain.txtsearchproduct_3.Text = ""
            frmMain.stocksdata.RecordSource = "select * from tblstocks"
            frmMain.stocksdata.Refresh
            frmMain.Enabled = True
            Unload Me
            frmMain.txtsearchproduct_3.SetFocus
            frmMain.cmdSearchCart.Caption = "CANCEL SEARCH"
            frmMain.cmdSearchCart.BackColor = &HC0C0FF
            MsgBox "Item Found!", vbInformation, ""
        End If
    End If
End If

If KeyAscii = 27 Then
    frmMain.Show
    frmMain.txtsearchproduct_3.Text = ""
    frmMain.stocksdata.RecordSource = "select * from tblstocks"
    frmMain.stocksdata.Refresh
    frmMain.transactdata.RecordSource = "select * from tbltransact"
    frmMain.transactdata.Refresh
    frmMain.Enabled = True
    Unload Me
    frmMain.txtsearchproduct_3.SetFocus
End If


If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789.- $%&*", Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

