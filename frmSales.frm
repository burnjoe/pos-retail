VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSalesRecord 
   BackColor       =   &H003E472C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   13815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmViewRecord 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.TextBox txtSearchItem 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   3975
      End
      Begin VB.CommandButton cmdClose1 
         BackColor       =   &H008080FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   13200
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
      Begin VB.Data productsalesdata 
         Caption         =   "SALES"
         Connect         =   "Access"
         DatabaseName    =   "database\database_sabana.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   8520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tblproductsales"
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdPrintReceipt 
         BackColor       =   &H00EAF1F4&
         Caption         =   "PRINT RECEIPT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6240
         Width           =   2295
      End
      Begin VB.Data printproductsalesdata 
         Caption         =   "FOR PRINT"
         Connect         =   "Access"
         DatabaseName    =   "database\database_sabana.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   10800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tblprintproductsales"
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmSales.frx":0000
         Height          =   4815
         Left            =   360
         OleObjectBlob   =   "frmSales.frx":0024
         TabIndex        =   1
         Top             =   1200
         Width           =   12855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH PRODUCT ID, NAME OR BARCODE:"
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
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM PURCHASED"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5760
         TabIndex        =   4
         Top             =   240
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmSalesRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose1_Click()
    frmMain.txtSearchOR.Text = ""
    frmMain.Enabled = True
    frmMain.Show
    frmMain.txtSearchOR.SetFocus
    Unload Me
End Sub

Private Sub cmdPrintReceipt_Click()
    txtSearchItem.Text = ""
    
    'PARA SA DATA REPORT TO
    productsalesprint.Sections("Section4").Controls("lblOR").Caption = frmMain.salesdata.Recordset.Fields("orno")
    productsalesprint.Sections("Section4").Controls("lblcashier").Caption = frmMain.salesdata.Recordset.Fields("cashier")
    productsalesprint.Sections("Section4").Controls("lbldate").Caption = frmMain.salesdata.Recordset.Fields("date")
    productsalesprint.Sections("Section4").Controls("lbltime").Caption = frmMain.salesdata.Recordset.Fields("time")
            
    productsalesprint.Sections("Section5").Controls("lbltotal").Caption = frmMain.salesdata.Recordset.Fields("total")
    productsalesprint.Sections("Section5").Controls("lblcash").Caption = frmMain.salesdata.Recordset.Fields("cash")
    productsalesprint.Sections("Section5").Controls("lblchange").Caption = frmMain.salesdata.Recordset.Fields("change")
    productsalesprint.Sections("Section5").Controls("lblquantity").Caption = frmMain.salesdata.Recordset.Fields("itempurchased")
    
    frmSalesRecord.Enabled = False
    
    Unload DataEnvironment1
    productsalesprint.Show
End Sub

Private Sub txtSearchItem_Change()
If txtSearchItem.Text = Empty Then
    printproductsalesdata.RecordSource = "select * from tblprintproductsales"
    printproductsalesdata.Refresh
Else
    printproductsalesdata.RecordSource = "select * from tblprintproductsales where productid= '" + txtSearchItem.Text + "'"
    printproductsalesdata.Refresh
    
    If printproductsalesdata.Recordset.EOF Then
        printproductsalesdata.RecordSource = "select * from tblprintproductsales where productname= '" + txtSearchItem.Text + "'"
        printproductsalesdata.Refresh
        
        If printproductsalesdata.Recordset.EOF Then
            printproductsalesdata.RecordSource = "select * from tblprintproductsales where barcode= '" + txtSearchItem.Text + "'"
            printproductsalesdata.Refresh
        End If
    End If
End If
End Sub

Private Sub txtSearchItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789()-_,.#!&%$* ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
