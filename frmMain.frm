VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H003E472C&
   BorderStyle     =   0  'None
   ClientHeight    =   9840
   ClientLeft      =   885
   ClientTop       =   540
   ClientWidth     =   18735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   18735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSideMenu 
      BackColor       =   &H004E8E6A&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   0
      TabIndex        =   33
      Top             =   240
      Width           =   2895
      Begin VB.Frame Frame2 
         BackColor       =   &H004E8E6A&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         TabIndex        =   63
         Top             =   1320
         Width           =   2415
         Begin VB.Label lbllast 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LAST NAME"
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
            Left            =   0
            TabIndex        =   65
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lblfirst 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FIRST NAME"
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
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   1185
         End
      End
      Begin VB.CommandButton cmdMyAccount 
         BackColor       =   &H0082D1B0&
         Caption         =   "MY ACCOUNT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Interval        =   5
         Left            =   2280
         Top             =   8280
      End
      Begin VB.CommandButton cmdLogOut 
         BackColor       =   &H0082D1B0&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6600
         Width           =   2655
      End
      Begin VB.CommandButton cmdManageAccounts 
         BackColor       =   &H0082D1B0&
         Caption         =   "MANAGE ACCOUNTS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5640
         Width           =   2655
      End
      Begin VB.CommandButton cmdInventory 
         BackColor       =   &H0082D1B0&
         Caption         =   "INVENTORY"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4680
         Width           =   2655
      End
      Begin VB.CommandButton cmdSales 
         BackColor       =   &H0082D1B0&
         Caption         =   "SALES"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         MaskColor       =   &H000080FF&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3720
         Width           =   2655
      End
      Begin VB.CommandButton cmdTransaction 
         BackColor       =   &H00DBF2E9&
         Caption         =   "TRANSACTION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Image imgUserType 
         Height          =   1080
         Left            =   240
         Picture         =   "frmMain.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label lblusertype 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADMIN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   37
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label lblDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day"
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
         Left            =   360
         TabIndex        =   36
         Top             =   9000
         Width           =   375
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   360
         TabIndex        =   35
         Top             =   8760
         Width           =   465
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   34
         Top             =   8280
         Width           =   840
      End
   End
   Begin VB.Frame frmSales 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   2880
      TabIndex        =   70
      Top             =   240
      Visible         =   0   'False
      Width           =   15855
      Begin VB.Data printsalesdata 
         Caption         =   "PRINT SALES"
         Connect         =   "Access"
         DatabaseName    =   "database\database_sabana.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   12360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tblprintsales"
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data salesdata 
         Caption         =   "SALES DATA"
         Connect         =   "Access"
         DatabaseName    =   "database\database_sabana.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   12360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tblsales"
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00EAF1F4&
         BorderStyle     =   0  'None
         Height          =   8535
         Left            =   0
         TabIndex        =   72
         Top             =   1080
         Width           =   15855
         Begin VB.CommandButton cmdSort 
            BackColor       =   &H0082D1B0&
            Caption         =   "SORT"
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
            Left            =   13920
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   1500
            Width           =   1215
         End
         Begin VB.CommandButton cmdDeleteRecord 
            BackColor       =   &H008080FF&
            Caption         =   "DELETE RECORD"
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
            Left            =   11760
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   6960
            Width           =   2295
         End
         Begin VB.CommandButton cmdPrintReport 
            BackColor       =   &H0082D1B0&
            Caption         =   "PRINT REPORT"
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
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   6960
            Width           =   2295
         End
         Begin VB.CommandButton cmdViewRecord 
            BackColor       =   &H0082D1B0&
            Caption         =   "VIEW RECORD"
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
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   6960
            Width           =   2295
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00EAF1F4&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            TabIndex        =   80
            Top             =   240
            Width           =   14775
            Begin VB.Label lbltotalsales 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   21.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   2160
               TabIndex        =   82
               Top             =   120
               Width           =   795
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL SALES:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   240
               TabIndex        =   81
               Top             =   240
               Width           =   1785
            End
         End
         Begin VB.TextBox txtSearchOR 
            BackColor       =   &H00D6D3D5&
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
            Left            =   720
            TabIndex        =   78
            Top             =   1560
            Width           =   3975
         End
         Begin MSDBGrid.DBGrid DBGrid2 
            Bindings        =   "frmMain.frx":CA73
            Height          =   4575
            Left            =   720
            OleObjectBlob   =   "frmMain.frx":CA8B
            TabIndex        =   73
            Top             =   2040
            Width           =   14415
         End
         Begin MSComCtl2.DTPicker dtsales 
            Height          =   330
            Left            =   10440
            TabIndex        =   86
            Top             =   1560
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarBackColor=   14078933
            CalendarTitleBackColor=   14078933
            CustomFormat    =   "MMMMdd, yyyy"
            Format          =   107020291
            CurrentDate     =   43905
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH BY DATE:"
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
            Left            =   10440
            TabIndex        =   88
            Top             =   1320
            Width           =   1350
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH OR #:"
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
            Left            =   720
            TabIndex        =   79
            Top             =   1320
            Width           =   1080
         End
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALES"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7635
         TabIndex        =   71
         Top             =   360
         Width           =   780
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   6840
         Picture         =   "frmMain.frx":DE6A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frmManageAccounts 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   2880
      TabIndex        =   29
      Top             =   240
      Width           =   15855
      Begin VB.Frame Frame4 
         BackColor       =   &H00EAF1F4&
         BorderStyle     =   0  'None
         Height          =   8535
         Left            =   0
         TabIndex        =   31
         Top             =   1080
         Width           =   15855
         Begin VB.CommandButton cmdView 
            BackColor       =   &H0082D1B0&
            Caption         =   "VIEW"
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
            Left            =   8880
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   6840
            Width           =   1935
         End
         Begin VB.TextBox txtsearchaccount 
            BackColor       =   &H00D6D3D5&
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
            Left            =   720
            TabIndex        =   18
            Top             =   1320
            Width           =   4935
         End
         Begin VB.ComboBox cmbAccountStatus 
            BackColor       =   &H00D6D3D5&
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
            ItemData        =   "frmMain.frx":1F4C3
            Left            =   13200
            List            =   "frmMain.frx":1F4D0
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1320
            Width           =   1935
         End
         Begin VB.ComboBox cmbUserType 
            BackColor       =   &H00D6D3D5&
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
            ItemData        =   "frmMain.frx":1F4F1
            Left            =   10440
            List            =   "frmMain.frx":1F4FE
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1320
            Width           =   2655
         End
         Begin VB.CommandButton cmdDeleteAdmin 
            BackColor       =   &H008080FF&
            Caption         =   "DELETE"
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
            Left            =   12720
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   6840
            Width           =   1935
         End
         Begin VB.CommandButton cmdUpdateAdmin 
            BackColor       =   &H0082D1B0&
            Caption         =   "UPDATE"
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
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   6840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddAdmin 
            BackColor       =   &H0082D1B0&
            Caption         =   "ADD ACCOUNT"
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
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   6840
            Width           =   1935
         End
         Begin VB.Data accountsdata 
            Caption         =   "ACCOUNTS"
            Connect         =   "Access"
            DatabaseName    =   "database/database_sabana.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   12000
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "tblaccounts"
            Top             =   240
            Visible         =   0   'False
            Width           =   2775
         End
         Begin MSDBGrid.DBGrid dbgridAccounts 
            Bindings        =   "frmMain.frx":1F518
            Height          =   4575
            Left            =   720
            OleObjectBlob   =   "frmMain.frx":1F533
            TabIndex        =   32
            Top             =   1800
            Width           =   14415
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SEARCH ACCOUNT ID:"
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
            Left            =   720
            TabIndex        =   41
            Top             =   1080
            Width           =   1710
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SORT BY STATUS:"
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
            Left            =   13200
            TabIndex        =   39
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SORT BY USER TYPE:"
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
            Left            =   10440
            TabIndex        =   38
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   5880
         Picture         =   "frmMain.frx":21635
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MANAGE ACCOUNTS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6675
         TabIndex        =   30
         Top             =   360
         Width           =   2745
      End
   End
   Begin VB.Frame frmInventory 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   15855
      Begin VB.CommandButton cmdStocks 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STOCKS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton cmdProducts 
         BackColor       =   &H00EAF1F4&
         Caption         =   "PRODUCTS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.Frame frmProducts 
         BackColor       =   &H00EAF1F4&
         BorderStyle     =   0  'None
         Height          =   8535
         Left            =   0
         TabIndex        =   28
         Top             =   1080
         Width           =   15855
         Begin VB.TextBox txtsearchproduct_2 
            BackColor       =   &H00D6D3D5&
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
            Left            =   720
            MaxLength       =   100
            TabIndex        =   6
            Top             =   1320
            Width           =   4935
         End
         Begin VB.CommandButton cmdDeleteProduct 
            BackColor       =   &H008080FF&
            Caption         =   "DELETE"
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
            Left            =   12720
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   6840
            Width           =   1935
         End
         Begin VB.CommandButton cmdViewProduct 
            BackColor       =   &H0082D1B0&
            Caption         =   "VIEW"
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
            Left            =   8880
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   6840
            Width           =   1935
         End
         Begin VB.CommandButton cmdUpdateProduct 
            BackColor       =   &H0082D1B0&
            Caption         =   "UPDATE"
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
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   6840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddProduct 
            BackColor       =   &H0082D1B0&
            Caption         =   "ADD PRODUCT"
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
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   6840
            Width           =   1935
         End
         Begin VB.Data productsdatainv 
            Caption         =   "For Products"
            Connect         =   "Access"
            DatabaseName    =   "database/database_sabana.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   12120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "tblproducts"
            Top             =   240
            Visible         =   0   'False
            Width           =   3060
         End
         Begin MSDBGrid.DBGrid DBGrid1 
            Bindings        =   "frmMain.frx":447F5
            Height          =   4575
            Left            =   720
            OleObjectBlob   =   "frmMain.frx":44813
            TabIndex        =   49
            Top             =   1800
            Width           =   14415
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT INFORMATION"
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
            Left            =   6480
            TabIndex        =   51
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label5 
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
            Left            =   720
            TabIndex        =   42
            Top             =   1080
            Width           =   3375
         End
      End
      Begin VB.Frame frmStocks 
         BackColor       =   &H00EAF1F4&
         BorderStyle     =   0  'None
         Height          =   8535
         Left            =   0
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   15855
         Begin VB.CommandButton Command1 
            BackColor       =   &H0082D1B0&
            Caption         =   "UPDATE STOCK QUANTITY"
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
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   6840
            Width           =   3015
         End
         Begin VB.CommandButton cmdDeleteStocks 
            BackColor       =   &H008080FF&
            Caption         =   "DELETE STOCKS"
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
            Left            =   11160
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   6840
            Width           =   2175
         End
         Begin VB.Data stocksdata 
            Caption         =   "TBLSTOCKS"
            Connect         =   "Access"
            DatabaseName    =   "database/database_sabana.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   11880
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "tblstocks"
            Top             =   240
            Visible         =   0   'False
            Width           =   3060
         End
         Begin VB.CommandButton cmdAddNewStock 
            BackColor       =   &H0082D1B0&
            Caption         =   "ADD NEW STOCK"
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
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   6840
            Width           =   2295
         End
         Begin VB.TextBox txtsearchproduct 
            BackColor       =   &H00D6D3D5&
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
            Left            =   1920
            MaxLength       =   30
            TabIndex        =   1
            Top             =   1320
            Width           =   4935
         End
         Begin MSDBGrid.DBGrid DBGrid5 
            Bindings        =   "frmMain.frx":456FE
            Height          =   4575
            Left            =   1920
            OleObjectBlob   =   "frmMain.frx":45717
            TabIndex        =   60
            Top             =   1800
            Width           =   12135
         End
         Begin VB.Label Label9 
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
            Left            =   1920
            TabIndex        =   61
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT STOCKS"
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
            Left            =   6840
            TabIndex        =   52
            Top             =   360
            Width           =   2115
         End
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   6720
         Picture         =   "frmMain.frx":467AA
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INVENTORY"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7440
         TabIndex        =   23
         Top             =   360
         Width           =   1545
      End
   End
   Begin VB.Frame frmTransaction 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   2880
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   15855
      Begin VB.Data transactdata 
         Caption         =   "TRANSACT"
         Connect         =   "Access"
         DatabaseName    =   "database\database_sabana.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   12000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "tbltransact"
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Frame frmCompleted 
         BackColor       =   &H003E472C&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   0
         TabIndex        =   74
         Top             =   3240
         Visible         =   0   'False
         Width           =   15855
         Begin VB.CommandButton cmdReprint 
            BackColor       =   &H0082D1B0&
            Caption         =   "REPRINT RECEIPT"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   1920
            Width           =   2535
         End
         Begin VB.CommandButton cmdNewTransaction 
            BackColor       =   &H0082D1B0&
            Caption         =   "NEW TRANSACTION"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   1920
            Width           =   2535
         End
         Begin VB.Image Image5 
            Height          =   735
            Left            =   7680
            Picture         =   "frmMain.frx":55E62
            Stretch         =   -1  'True
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSACTION COMPLETED"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   585
            Left            =   5325
            TabIndex        =   75
            Top             =   1080
            Width           =   5490
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EAF1F4&
         BorderStyle     =   0  'None
         Height          =   8535
         Left            =   0
         TabIndex        =   27
         Top             =   1080
         Width           =   15855
         Begin VB.CommandButton cmdAddToCart 
            BackColor       =   &H0082D1B0&
            Caption         =   "ADD TO CART"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8160
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtsearchproduct_3 
            BackColor       =   &H00D6D3D5&
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
            Left            =   240
            TabIndex        =   53
            Top             =   480
            Width           =   7815
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            Height          =   8055
            Left            =   10080
            TabIndex        =   43
            Top             =   240
            Width           =   5535
            Begin VB.Frame Frame6 
               BackColor       =   &H003E472C&
               BorderStyle     =   0  'None
               Height          =   1455
               Left            =   120
               TabIndex        =   54
               Top             =   120
               Width           =   5295
               Begin VB.Frame Frame7 
                  BackColor       =   &H003E472C&
                  BorderStyle     =   0  'None
                  Height          =   495
                  Left            =   1800
                  TabIndex        =   68
                  Top             =   840
                  Width           =   3375
                  Begin VB.Label lblchange 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "0.00"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   15.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   390
                     Left            =   2580
                     TabIndex        =   69
                     Top             =   0
                     Width           =   585
                  End
               End
               Begin VB.Frame Frame5 
                  BackColor       =   &H003E472C&
                  BorderStyle     =   0  'None
                  Height          =   615
                  Left            =   1800
                  TabIndex        =   66
                  Top             =   120
                  Width           =   3375
                  Begin VB.Label lbltotal 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "0.00"
                     BeginProperty Font 
                        Name            =   "Calibri"
                        Size            =   24
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   585
                     Left            =   2385
                     TabIndex        =   67
                     Top             =   0
                     Width           =   855
                  End
               End
               Begin VB.Label label0 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "TOTAL:"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   24
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   585
                  Left            =   240
                  TabIndex        =   56
                  Top             =   120
                  Width           =   1440
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "CHANGE:"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   390
                  Left            =   240
                  TabIndex        =   55
                  Top             =   840
                  Width           =   1200
               End
            End
            Begin VB.CommandButton cmdEditQty 
               BackColor       =   &H0082D1B0&
               Caption         =   "ADD / DEDUCT QUANTITY"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   5160
               Width           =   2655
            End
            Begin VB.CommandButton cmdRemoveItem 
               BackColor       =   &H00C0C0FF&
               Caption         =   "REMOVE ITEM"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2880
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   5160
               Width           =   2535
            End
            Begin VB.CommandButton cmdSearchCart 
               BackColor       =   &H0082D1B0&
               Caption         =   "SEARCH ITEM FROM CART"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   46
               Top             =   6000
               Width           =   2655
            End
            Begin VB.CommandButton cmdCancelTransaction 
               BackColor       =   &H00C0C0FF&
               Caption         =   "CANCEL TRANSACTION"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   2880
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   6000
               Width           =   2535
            End
            Begin VB.CommandButton cmdPay 
               BackColor       =   &H0080FF80&
               Caption         =   "PAY"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   6840
               Width           =   5295
            End
            Begin MSDBGrid.DBGrid DBGrid4 
               Bindings        =   "frmMain.frx":76DF9
               Height          =   3375
               Left            =   120
               OleObjectBlob   =   "frmMain.frx":76E14
               TabIndex        =   59
               Top             =   1680
               Width           =   5295
            End
         End
         Begin MSDBGrid.DBGrid DBGrid3 
            Bindings        =   "frmMain.frx":77CFF
            Height          =   7335
            Left            =   240
            OleObjectBlob   =   "frmMain.frx":77D18
            TabIndex        =   58
            Top             =   960
            Width           =   9615
         End
         Begin VB.Label Label2 
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
            Left            =   240
            TabIndex        =   62
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   6600
         Picture         =   "frmMain.frx":78DB7
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSACTION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7395
         TabIndex        =   26
         Top             =   360
         Width           =   1905
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public menubutton As String
Public idno, sortdate As String
Public totalsales As Currency

Private Sub cmbAccountStatus_Click()
If cmbAccountStatus.Text = "GENERAL" And cmbUserType.Text = "GENERAL" Then
    If Not txtsearchaccount.Text = Empty Then
        accountsdata.RecordSource = "select * from tblaccounts where accid= '" + txtsearchaccount.Text + "'"
        accountsdata.Refresh
    Else
        accountsdata.RecordSource = "select * from tblaccounts"
        accountsdata.Refresh
    End If
Else
    If cmbAccountStatus.Text = "ACTIVE" And cmbUserType.Text = "GENERAL" Then
        If Not txtsearchaccount.Text = Empty Then
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and accid= '" + txtsearchaccount + "'"
            accountsdata.Refresh
        Else
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "'"
            accountsdata.Refresh
        End If
    Else
        If cmbAccountStatus.Text = "NOT ACTIVE" And cmbUserType.Text = "GENERAL" Then
            If Not txtsearchaccount.Text = Empty Then
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and accid= '" + txtsearchaccount.Text + "'"
                accountsdata.Refresh
            Else
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "'"
                accountsdata.Refresh
            End If
        End If
    End If
End If


If cmbAccountStatus.Text = "GENERAL" And cmbUserType.Text = "ADMIN" Then
    If Not txtsearchaccount.Text = Empty Then
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount + "'"
        accountsdata.Refresh
    Else
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "'"
        accountsdata.Refresh
    End If
Else
    If cmbAccountStatus.Text = "ACTIVE" And cmbUserType.Text = "ADMIN" Then
        If Not txtsearchaccount.Text = Empty Then
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
            accountsdata.Refresh
        Else
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
            accountsdata.Refresh
        End If
    Else
        If cmbAccountStatus.Text = "NOT ACTIVE" And cmbUserType.Text = "ADMIN" Then
            If Not txtsearchaccount.Text = Empty Then
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
                accountsdata.Refresh
            Else
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
                accountsdata.Refresh
            End If
        End If
    End If
End If


If cmbAccountStatus.Text = "GENERAL" And cmbUserType.Text = "USER" Then
    If Not txtsearchaccount.Text = Empty Then
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
        accountsdata.Refresh
    Else
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "'"
        accountsdata.Refresh
    End If
Else
    If cmbAccountStatus.Text = "ACTIVE" And cmbUserType.Text = "USER" Then
        If Not txtsearchaccount.Text = Empty Then
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
            accountsdata.Refresh
        Else
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
            accountsdata.Refresh
        End If
    Else
        If cmbAccountStatus.Text = "NOT ACTIVE" And cmbUserType.Text = "USER" Then
            If Not txtsearchaccount.Text = Empty Then
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
                accountsdata.Refresh
            Else
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
                accountsdata.Refresh
            End If
        End If
    End If
End If
End Sub

Private Sub cmbUserType_Click()
If cmbAccountStatus.Text = "GENERAL" And cmbUserType.Text = "GENERAL" Then
    If Not txtsearchaccount.Text = Empty Then
        accountsdata.RecordSource = "select * from tblaccounts where accid= '" + txtsearchaccount.Text + "'"
        accountsdata.Refresh
    Else
        accountsdata.RecordSource = "select * from tblaccounts"
        accountsdata.Refresh
    End If
Else
    If cmbAccountStatus.Text = "ACTIVE" And cmbUserType.Text = "GENERAL" Then
        If Not txtsearchaccount.Text = Empty Then
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and accid= '" + txtsearchaccount + "'"
            accountsdata.Refresh
        Else
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "'"
            accountsdata.Refresh
        End If
    Else
        If cmbAccountStatus.Text = "NOT ACTIVE" And cmbUserType.Text = "GENERAL" Then
            If Not txtsearchaccount.Text = Empty Then
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and accid= '" + txtsearchaccount.Text + "'"
                accountsdata.Refresh
            Else
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "'"
                accountsdata.Refresh
            End If
        End If
    End If
End If


If cmbAccountStatus.Text = "GENERAL" And cmbUserType.Text = "ADMIN" Then
    If Not txtsearchaccount.Text = Empty Then
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount + "'"
        accountsdata.Refresh
    Else
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "'"
        accountsdata.Refresh
    End If
Else
    If cmbAccountStatus.Text = "ACTIVE" And cmbUserType.Text = "ADMIN" Then
        If Not txtsearchaccount.Text = Empty Then
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
            accountsdata.Refresh
        Else
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
            accountsdata.Refresh
        End If
    Else
        If cmbAccountStatus.Text = "NOT ACTIVE" And cmbUserType.Text = "ADMIN" Then
            If Not txtsearchaccount.Text = Empty Then
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
                accountsdata.Refresh
            Else
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
                accountsdata.Refresh
            End If
        End If
    End If
End If


If cmbAccountStatus.Text = "GENERAL" And cmbUserType.Text = "USER" Then
    If Not txtsearchaccount.Text = Empty Then
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
        accountsdata.Refresh
    Else
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "'"
        accountsdata.Refresh
    End If
Else
    If cmbAccountStatus.Text = "ACTIVE" And cmbUserType.Text = "USER" Then
        If Not txtsearchaccount.Text = Empty Then
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
            accountsdata.Refresh
        Else
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
            accountsdata.Refresh
        End If
    Else
        If cmbAccountStatus.Text = "NOT ACTIVE" And cmbUserType.Text = "USER" Then
            If Not txtsearchaccount.Text = Empty Then
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
                accountsdata.Refresh
            Else
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
                accountsdata.Refresh
            End If
        End If
    End If
End If
End Sub


Private Sub cmdAddAdmin_Click() ' kailangan
frmAskAdminPassword.whatbutton = "addaccount"
frmAskAdminPassword.Show
frmMain.Enabled = False
frmAskAdminPassword.txtAskAdminPass.SetFocus
cmdAddAdmin.BackColor = &HDBF2E9
End Sub

Private Sub cmdCategories_Click() 'kailangan

frmCategories.Visible = True
frmProducts.Visible = False
cmdProducts.BackColor = &H82D1B0
cmdCategories.BackColor = &HDBF2E9

End Sub



Private Sub cmdAddNewStock_Click()
    frmAskAdminPassword.Show
    frmAskAdminPassword.whatbutton = "addnewstock"
    Me.Enabled = False
End Sub

Private Sub cmdAddToCart_Click()
If stocksdata.Recordset.EOF Then
    MsgBox "Select Item To Add!", vbCritical, ""
    txtsearchproduct_3.Text = ""
    txtsearchproduct_3.SetFocus
Else
    If stocksdata.Recordset.Fields("stocks") = 0 Then
        MsgBox "Out Of Stocks!!", vbCritical, ""
        frmMain.txtsearchproduct_3.SetFocus
        txtsearchproduct_3.Text = ""
    Else
        frmPopUps.Show
        frmPopUps.frmAddQty.Visible = True
        Me.Enabled = False
    End If
End If
End Sub

Private Sub cmdCancelTransaction_Click()
Dim answer As String

' may kulang pa
If transactdata.Recordset.EOF Then
    frmMain.frmSideMenu.Enabled = True
    frmMain.cmdMyAccount.BackColor = &H82D1B0
    frmMain.cmdTransaction.BackColor = &HDBF2E9
    frmMain.cmdSales.BackColor = &H82D1B0
    frmMain.cmdInventory.BackColor = &H82D1B0
    frmMain.cmdManageAccounts.BackColor = &H82D1B0
    frmMain.cmdLogOut.BackColor = &H82D1B0
    txtsearchproduct_3.SetFocus
    txtsearchproduct_3.Text = ""
    lbltotal.Caption = "0.00"
    frmPopUps.total = 0
    lblchange.Caption = "0.00"
    lbltotal.FontSize = 24
Else
    answer = MsgBox("Cancel Current Transaction?", vbQuestion + vbYesNo, "")
      
    If answer = vbYes Then
        Do
        stocksdata.RecordSource = "select * from tblstocks where barcode= '" + transactdata.Recordset.Fields("barcode") + "'"
        stocksdata.Refresh
        
        If Not stocksdata.Recordset.EOF Then
            stocksdata.Recordset.Edit
            stocksdata.Recordset.Fields("stocks") = Val(stocksdata.Recordset.Fields("stocks")) + Val(transactdata.Recordset.Fields("quantity"))
            stocksdata.Recordset.Update
            transactdata.Recordset.Delete
        End If
            
        transactdata.Refresh
        Loop Until transactdata.Recordset.EOF
        
        stocksdata.RecordSource = "select * from tblstocks"
        stocksdata.Refresh
        transactdata.RecordSource = "select * from tbltransact"
        transactdata.Refresh
        
        lbltotal.Caption = "0.00"
        lblchange.Caption = "0.00"
        frmPopUps.total = 0
        
        frmMain.frmSideMenu.Enabled = True
        frmMain.cmdMyAccount.BackColor = &H82D1B0
        frmMain.cmdTransaction.BackColor = &HDBF2E9
        frmMain.cmdSales.BackColor = &H82D1B0
        frmMain.cmdInventory.BackColor = &H82D1B0
        frmMain.cmdManageAccounts.BackColor = &H82D1B0
        frmMain.cmdLogOut.BackColor = &H82D1B0
        txtsearchproduct_3.SetFocus
        txtsearchproduct_3.Text = ""
        lbltotal.FontSize = 24
    End If
End If
End Sub

Private Sub cmdDeleteAdmin_Click()
cmdDeleteAdmin.BackColor = &HC0C0FF

If accountsdata.Recordset.EOF Then
    MsgBox "Select Account To Delete!", vbCritical, ""
    txtsearchaccount.Text = ""
    txtsearchaccount.SetFocus
    cmdDeleteAdmin.BackColor = &H8080FF
    cmbUserType = "GENERAL"
    cmbAccountStatus = "GENERAL"
Else
    frmAskAdminPassword.whatbutton = "deleteaccount"
    Me.Enabled = False
    frmAskAdminPassword.Show
End If
End Sub

Private Sub cmdDeleteProduct_Click()
If productsdatainv.Recordset.EOF Then
    MsgBox "Select Product To Delete!", vbCritical, ""
    txtsearchproduct_2.Text = ""
    txtsearchproduct_2.SetFocus
Else
    If productsdatainv.Recordset.EOF Then
        MsgBox "Select Product To Delete!", vbCritical, ""
        txtsearchproduct_2.Text = ""
        txtsearchproduct_2.SetFocus
    Else
        frmAskAdminPassword.whatbutton = "deleteproduct"
        frmAskAdminPassword.Show
        frmMain.Enabled = False
        frmAskAdminPassword.txtAskAdminPass.SetFocus
        idno = productsdatainv.Recordset.Fields("productid")
    End If
End If
End Sub

Private Sub cmdDeleteRecord_Click()
If salesdata.Recordset.EOF Then
    MsgBox "Select Record To Delete!", vbCritical, ""
    txtSearchOR.Text = ""
    txtSearchOR.SetFocus
    salesdata.Refresh
Else
    Dim replydel As String
        
    replydel = MsgBox("Delete This Record?  ( OR #: " & salesdata.Recordset.Fields("orno") & " - " & salesdata.Recordset.Fields("date") & " )", vbQuestion + vbYesNo, "")
        
    If replydel = vbYes Then
        MsgBox "Deleted Record!", vbInformation, ""
     
        printsalesdata.RecordSource = "select * from tblprintsales where orno= '" + salesdata.Recordset.Fields("orno") + "' and date= '" + salesdata.Recordset.Fields("date") + "'"
        printsalesdata.Refresh
            
        If Not printsalesdata.Recordset.EOF Then
            printsalesdata.Recordset.Delete
        End If
            
        frmSalesRecord.productsalesdata.RecordSource = "select * from tblproductsales where orno= '" + salesdata.Recordset.Fields("orno") + "'"
        frmSalesRecord.productsalesdata.Refresh
            
        If Not frmSalesRecord.productsalesdata.Recordset.EOF Then
            Do
            frmSalesRecord.productsalesdata.Recordset.Delete
            frmSalesRecord.productsalesdata.RecordSource = "select * from tblproductsales where orno= '" + salesdata.Recordset.Fields("orno") + "'"
            frmSalesRecord.productsalesdata.Refresh
            Loop Until frmSalesRecord.productsalesdata.Recordset.EOF
        End If
            
        salesdata.Recordset.Delete
            
        totalsales = 0
            
        If cmdSort.Caption = "CANCEL" Then
            salesdata.RecordSource = "select * from tblsales where date= '" + sortdate + "'"
            salesdata.Refresh
                
            If Not salesdata.Recordset.EOF Then
                Do
                salesdata.RecordSource = "select * from tblsales where date= '" + sortdate + "'"
                salesdata.Refresh
                totalsales = totalsales + FormatNumber(frmMain.salesdata.Recordset.Fields("total"))
                frmMain.salesdata.Recordset.MoveNext
                Loop Until frmMain.salesdata.Recordset.EOF
            Else
                sortdate = 0
                    
                salesdata.RecordSource = "select * from tblsales"
                salesdata.Refresh
                    
                    
                    
                If Not salesdata.Recordset.EOF Then
                    Do
                    totalsales = totalsales + FormatNumber(frmMain.salesdata.Recordset.Fields("total"))
                    frmMain.salesdata.Recordset.MoveNext
                    Loop Until frmMain.salesdata.Recordset.EOF
                        
                    lbltotalsales.Caption = FormatNumber(totalsales)
                Else
                    lbltotalsales.Caption = "0.00"
                End If
                    
                    
                    
                salesdata.Refresh
                cmdSort.Caption = "SORT"
                cmdSort.BackColor = &H82D1B0
                        
            End If
                    
                lbltotalsales.Caption = FormatNumber(totalsales)
                salesdata.Refresh
        Else
            sortdate = 0
                        
            salesdata.RecordSource = "select * from tblsales"
            salesdata.Refresh
                        
                        
            If Not salesdata.Recordset.EOF Then
                Do
                totalsales = totalsales + FormatNumber(frmMain.salesdata.Recordset.Fields("total"))
                frmMain.salesdata.Recordset.MoveNext
                Loop Until frmMain.salesdata.Recordset.EOF
                            
                lbltotalsales.Caption = FormatNumber(totalsales)
            Else
                lbltotalsales.Caption = "0.00"
            End If
            
            salesdata.Refresh
            cmdSort.Caption = "SORT"
            cmdSort.BackColor = &H82D1B0
        End If
        lbltotalsales.Caption = FormatNumber(totalsales)
        salesdata.Refresh
        txtSearchOR.Text = ""
        txtSearchOR.SetFocus
    End If
End If
End Sub

Private Sub cmdDeleteStocks_Click()
If stocksdata.Recordset.EOF Then
    MsgBox "Select Stock To Delete!", vbCritical, ""
    txtsearchproduct.Text = ""
    txtsearchproduct.SetFocus
Else
    frmAskAdminPassword.whatbutton = "deletestock"
    frmAskAdminPassword.Show
    Me.Enabled = False
    frmAskAdminPassword.txtAskAdminPass.SetFocus
End If
End Sub

Private Sub cmdEditQty_Click()
Dim reply11 As String
If transactdata.Recordset.EOF Then
    MsgBox "Select Item To Add Or Deduct Quantity!", vbCritical, ""
    txtsearchproduct_3.SetFocus
    txtsearchproduct_3.Text = ""
Else
    reply11 = MsgBox("Add Or Deduct Quantity Of This Product (" & transactdata.Recordset.Fields("barcode") & " - " & transactdata.Recordset.Fields("productname") & ") ?", vbQuestion + vbYesNo, "")
    
    If reply11 = vbYes Then
        frmMain.Enabled = False
        frmPopUps.Show
        frmPopUps.frmEditQty.Visible = True
        frmPopUps.cmbEditType.Text = "ADD"
        frmPopUps.lbledittype.Caption = "ADD QUANTITY:"
    End If
End If
End Sub

Private Sub cmdInventory_Click() 'kailangan

productsdatainv.RecordSource = "select * from tblproducts"
productsdatainv.Refresh

printsalesdata.RecordSource = "select * from tblprintsales"
printsalesdata.Refresh

If Not printsalesdata.Recordset.EOF Then
    Do
    frmMain.printsalesdata.Recordset.Delete
    printsalesdata.Refresh
    Loop Until printsalesdata.Recordset.EOF
End If

cmdSort.Caption = "SORT"
cmdSort.BackColor = &H82D1B0

cmdInventory.BackColor = &HDBF2E9
cmdTransaction.BackColor = &H82D1B0
cmdSales.BackColor = &H82D1B0
cmdManageAccounts.BackColor = &H82D1B0
cmdLogOut.BackColor = &H82D1B0
frmInventory.Visible = True
frmTransaction.Visible = False
frmManageAccounts.Visible = False
frmSales.Visible = False
frmPopUps.total = 0

If menubutton = "inventoryadmin" Then
    frmProducts.Visible = True
    frmStocks.Visible = False
    txtsearchproduct_2.SetFocus
    txtsearchproduct_2.Text = ""
    txtsearchproduct.Text = ""
    cmdProducts.BackColor = &HEAF1F4
    cmdStocks.BackColor = &HC0C0C0
Else
    frmProducts.Visible = False
    frmStocks.Visible = True
    txtsearchproduct.SetFocus
    txtsearchproduct.Text = ""
    cmdProducts.Visible = False
    cmdStocks.Visible = False
End If



'MULA DITO - PARA SA TRANSACTION REFRESH
'PARA SA TRANSACTION GUI
    frmMain.frmCompleted.Visible = False
    frmMain.Frame1.Enabled = True
    frmMain.frmCompleted.Enabled = False
    frmMain.lbltotal.FontSize = 24
    
'COLORS
    frmMain.cmdAddToCart.BackColor = &H82D1B0
    frmMain.cmdEditQty.BackColor = &H82D1B0
    frmMain.cmdRemoveItem.BackColor = &HC0C0FF
    frmMain.cmdSearchCart.BackColor = &H82D1B0
    frmMain.cmdCancelTransaction.BackColor = &HC0C0FF
    frmMain.cmdPay.BackColor = &H80FF80
    frmMain.DBGrid3.BackColor = &H80000005
    frmMain.DBGrid4.BackColor = &H80000005
    frmMain.frmTransaction.BackColor = &H82D1B0
    frmMain.Frame6.BackColor = &H3E472C
    frmMain.Frame5.BackColor = &H3E472C
    frmMain.Frame7.BackColor = &H3E472C
    frmMain.Image3.Picture = LoadPicture(App.Path & "/images/cart_logo_2.jpg")
    
'NEW TRANSACTION AUTO
    transactdata.RecordSource = "select * from tbltransact"
    transactdata.Refresh
    
    If Not transactdata.Recordset.EOF Then
        Do
        transactdata.Recordset.Delete
        transactdata.Refresh
        Loop Until transactdata.Recordset.EOF
    End If

'DATA REFRESHES
    stocksdata.RecordSource = "select * from tblstocks"
    stocksdata.Refresh
    transactdata.RecordSource = "select * from tbltransact"
    transactdata.Refresh
    
'LABEL REFRESHES
    lbltotal.Caption = "0.00"
    lblchange.Caption = "0.00"
'HANGGANG DITO LANG


End Sub

Private Sub cmdLogOut_Click() 'kailangan

cmdLogOut.BackColor = &HDBF2E9


frmExitType.Show
Me.Enabled = False

End Sub

Private Sub cmdManageAccounts_Click() 'kailangan

If cmdManageAccounts.Caption = "MANAGE ACCOUNTS" Then
    cmdManageAccounts.BackColor = &HDBF2E9
    cmdSales.BackColor = &H82D1B0
    cmdInventory.BackColor = &H82D1B0
    cmdTransaction.BackColor = &H82D1B0
    cmdLogOut.BackColor = &H82D1B0
    frmManageAccounts.Visible = True
    frmTransaction.Visible = False
    frmInventory.Visible = False
    frmSales.Visible = False
    frmPopUps.total = 0
    
    printsalesdata.RecordSource = "select * from tblprintsales"
    printsalesdata.Refresh
    
    If Not printsalesdata.Recordset.EOF Then
        Do
        frmMain.printsalesdata.Recordset.Delete
        printsalesdata.Refresh
        Loop Until printsalesdata.Recordset.EOF
    End If
    
    cmdSort.Caption = "SORT"
    cmdSort.BackColor = &H82D1B0
   
    'laman niya
    cmbUserType = "GENERAL"
    cmbAccountStatus = "GENERAL"
    txtsearchaccount = ""
    txtsearchaccount.SetFocus
    
'MULA DITO - PARA SA TRANSACTION REFRESH
'PARA SA TRANSACTION GUI
    frmMain.frmCompleted.Visible = False
    frmMain.Frame1.Enabled = True
    frmMain.frmCompleted.Enabled = False
    frmMain.lbltotal.FontSize = 24
    
'COLORS
    frmMain.cmdAddToCart.BackColor = &H82D1B0
    frmMain.cmdEditQty.BackColor = &H82D1B0
    frmMain.cmdRemoveItem.BackColor = &HC0C0FF
    frmMain.cmdSearchCart.BackColor = &H82D1B0
    frmMain.cmdCancelTransaction.BackColor = &HC0C0FF
    frmMain.cmdPay.BackColor = &H80FF80
    frmMain.DBGrid3.BackColor = &H80000005
    frmMain.DBGrid4.BackColor = &H80000005
    frmMain.frmTransaction.BackColor = &H82D1B0
    frmMain.Frame6.BackColor = &H3E472C
    frmMain.Frame5.BackColor = &H3E472C
    frmMain.Frame7.BackColor = &H3E472C
    frmMain.Image3.Picture = LoadPicture(App.Path & "/images/cart_logo_2.jpg")
    
'NEW TRANSACTION AUTO
    transactdata.RecordSource = "select * from tbltransact"
    transactdata.Refresh
    
    If Not transactdata.Recordset.EOF Then
        Do
        transactdata.Recordset.Delete
        transactdata.Refresh
        Loop Until transactdata.Recordset.EOF
    End If

'DATA REFRESHES
    stocksdata.RecordSource = "select * from tblstocks"
    stocksdata.Refresh
    transactdata.RecordSource = "select * from tbltransact"
    transactdata.Refresh
    
'LABEL REFRESHES
    lbltotal.Caption = "0.00"
    lblchange.Caption = "0.00"
'HANGGANG DITO LANG


    
End If

If cmdManageAccounts.Caption = "EXIT" Then
    cmdManageAccounts.BackColor = &HDBF2E9
    
    frmExitType.Show
    Me.Enabled = False
End If


End Sub

Private Sub cmdMyAccount_Click()

frmAskAdminPassword.whatbutton = "myaccount"
cmdMyAccount.BackColor = &HDBF2E9
frmMain.Enabled = False
frmAskAdminPassword.Show
frmAskAdminPassword.txtAskAdminPass.SetFocus

End Sub

Private Sub cmdNewTransaction_Click()
frmMain.frmCompleted.Visible = False
frmMain.Frame1.Enabled = True
frmMain.frmCompleted.Enabled = False
frmMain.lbltotal.FontSize = 24
    
'COLORS
frmMain.cmdAddToCart.BackColor = &H82D1B0
frmMain.cmdEditQty.BackColor = &H82D1B0
frmMain.cmdRemoveItem.BackColor = &HC0C0FF
frmMain.cmdSearchCart.BackColor = &H82D1B0
frmMain.cmdCancelTransaction.BackColor = &HC0C0FF
frmMain.cmdPay.BackColor = &H80FF80
frmMain.DBGrid3.BackColor = &H80000005
frmMain.DBGrid4.BackColor = &H80000005
frmMain.Frame6.BackColor = &H3E472C
frmMain.Frame5.BackColor = &H3E472C
frmMain.Frame7.BackColor = &H3E472C
frmMain.frmTransaction.BackColor = &H82D1B0
frmMain.Image3.Picture = LoadPicture(App.Path & "/images/cart_logo_2.jpg")

'DELETING PREVIOUS RECORDS
transactdata.RecordSource = "select * from tbltransact"
transactdata.Refresh

Do
transactdata.Recordset.Delete
transactdata.Refresh
Loop Until transactdata.Recordset.EOF

frmPopUps.total = 0
frmMain.lbltotal = "0.00"
frmMain.lblchange = "0.00"

txtsearchproduct_3.SetFocus
End Sub

Private Sub cmdPay_Click()
If transactdata.Recordset.EOF Then
    MsgBox "Cart Is Empty!", vbCritical, ""
    txtsearchproduct_3.SetFocus
    txtsearchproduct_3.Text = ""
Else
    frmMain.Enabled = False
    frmPopUps.Show
    frmPopUps.frmPay.Visible = True
    cmdSearchCart.Caption = "SEARCH ITEM FROM CART"
    cmdSearchCart.BackColor = &H82D1B0
    transactdata.RecordSource = "select * from tbltransact"
    transactdata.Refresh
End If
End Sub

Private Sub cmdPrintReport_Click()
salesdata.Refresh

If salesdata.Recordset.EOF Then
    MsgBox "No Records!", vbCritical, ""
    txtSearchOR.Text = ""
    txtSearchOR.SetFocus
Else
    If cmdSort.Caption = "CANCEL" Then
        salesreportprint.Sections("Section2").Controls("lbldate1").Visible = True
        salesreportprint.Sections("Section2").Controls("lbldate2").Caption = frmMain.sortdate
        salesreportprint.Sections("Section2").Controls("lbltotalsales").Caption = frmMain.lbltotalsales.Caption
        frmMain.Enabled = False
        Unload DataEnvironment1
        salesreportprint.Show
    Else
        allsalesprint.Sections("Section2").Controls("lbldate1").Visible = False
        allsalesprint.Sections("Section2").Controls("lbldate2").Caption = "ALL RECORDS"
        allsalesprint.Sections("Section2").Controls("lbltotalsales").Caption = frmMain.lbltotalsales.Caption
        frmMain.Enabled = False
        Unload DataEnvironment1
        allsalesprint.Show
    End If
End If
End Sub

Private Sub cmdProducts_Click() 'kailangan

frmProducts.Visible = True
frmStocks.Visible = False
cmdProducts.BackColor = &HEAF1F4
cmdStocks.BackColor = &HC0C0C0
txtsearchproduct_2.SetFocus
txtsearchproduct_2.Text = ""
productsdatainv.RecordSource = "select * from tblproducts"
productsdatainv.Refresh

End Sub

Private Sub cmdRemoveItem_Click()
Dim reply9 As String

If transactdata.Recordset.EOF Then
    MsgBox "Select Item To Remove!", vbCritical, ""
    txtsearchproduct_3.SetFocus
    txtsearchproduct_3.Text = ""
Else
    reply9 = MsgBox("Remove This Item  (" & transactdata.Recordset.Fields("barcode") & " - " & transactdata.Recordset.Fields("productname") & ")  ?", vbQuestion + vbYesNo, "")
    
    If reply9 = vbYes Then
        frmPopUps.total = Val(frmPopUps.total) - FormatNumber(transactdata.Recordset.Fields("price"))
        lbltotal.Caption = FormatNumber(frmPopUps.total)
        
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
        
        stocksdata.RecordSource = "select * from tblstocks where barcode= '" + transactdata.Recordset.Fields("barcode") + "'"
        stocksdata.Refresh
        
        If Not stocksdata.Recordset.EOF Then
            stocksdata.Recordset.Edit
            stocksdata.Recordset.Fields("stocks") = Val(stocksdata.Recordset.Fields("stocks")) + Val(transactdata.Recordset.Fields("quantity"))
            stocksdata.Recordset.Update
            
            stocksdata.RecordSource = "select * from tblstocks"
            stocksdata.Refresh
        End If
   
        transactdata.Recordset.Delete
        
        txtsearchproduct_3.SetFocus
        
        
    End If
    transactdata.RecordSource = "select * from tbltransact"
    transactdata.Refresh
    
    'PARANG CANCEL TRANSACTION
    If transactdata.Recordset.EOF Then
        frmMain.frmSideMenu.Enabled = True
        frmMain.cmdMyAccount.BackColor = &H82D1B0
        frmMain.cmdTransaction.BackColor = &HDBF2E9
        frmMain.cmdSales.BackColor = &H82D1B0
        frmMain.cmdInventory.BackColor = &H82D1B0
        frmMain.cmdManageAccounts.BackColor = &H82D1B0
        frmMain.cmdLogOut.BackColor = &H82D1B0
        txtsearchproduct_3.SetFocus
        txtsearchproduct_3.Text = ""
        lbltotal.Caption = "0.00"
        frmPopUps.total = 0
        lblchange.Caption = "0.00"
        lbltotal.FontSize = 24
    End If
    
    cmdSearchCart.Caption = "SEARCH ITEM FROM CART"
    cmdSearchCart.BackColor = &H82D1B0
End If
End Sub

Private Sub cmdReprint_Click()
frmMain.Enabled = False

transactprint.Sections("Section4").Controls("lblOR").Caption = frmPopUps.ornumber
transactprint.Sections("Section4").Controls("lblcashier").Caption = frmLogin.accountsdata.Recordset.Fields("accid") & "  " & frmLogin.accountsdata.Recordset.Fields("first")
transactprint.Sections("Section4").Controls("lbldate").Caption = frmPopUps.date1
transactprint.Sections("Section4").Controls("lbltime").Caption = frmPopUps.time1

transactprint.Sections("Section5").Controls("lbltotal").Caption = Format(Val(frmMain.lbltotal.Caption), "0.00")
transactprint.Sections("Section5").Controls("lblcash").Caption = Format(frmPopUps.pay, "0.00")
transactprint.Sections("Section5").Controls("lblchange").Caption = Format(Val(frmMain.lblchange.Caption), "0.00")
transactprint.Sections("Section5").Controls("lblquantity").Caption = frmPopUps.qty

Unload DataEnvironment1
transactprint.Show
End Sub

Private Sub cmdSearchCart_Click()
If transactdata.Recordset.EOF Then
    MsgBox "Cart Is Empty!", vbCritical, ""
    txtsearchproduct_3.SetFocus
    txtsearchproduct_3.Text = ""
Else
    If cmdSearchCart.Caption = "CANCEL SEARCH" Then
        transactdata.RecordSource = "select * from tbltransact"
        transactdata.Refresh
        cmdSearchCart.Caption = "SEARCH ITEM FROM CART"
        cmdSearchCart.BackColor = &H82D1B0
        txtsearchproduct_3.SetFocus
    Else
        frmMain.Enabled = False
        frmPopUps.Show
        frmPopUps.frmSearchCart.Visible = True
    End If
End If
End Sub

Private Sub cmdSort_Click()

txtSearchOR.Text = ""

If cmdSort.Caption = "SORT" Then
    totalsales = 0
    
    sortdate = Format(dtsales, "mmmm dd, yyyy")
    
    salesdata.RecordSource = "select * from tblsales where date= '" + sortdate + "'"
    salesdata.Refresh
     
    If Not salesdata.Recordset.EOF Then
        MsgBox "Record Sorted!", vbInformation, ""
        Do
        totalsales = totalsales + FormatNumber(frmMain.salesdata.Recordset.Fields("total"))
        frmMain.printsalesdata.Recordset.AddNew
        frmMain.printsalesdata.Recordset.Fields("orno") = salesdata.Recordset.Fields("orno")
        frmMain.printsalesdata.Recordset.Fields("cashier") = salesdata.Recordset.Fields("cashier")
        frmMain.printsalesdata.Recordset.Fields("date") = salesdata.Recordset.Fields("date")
        frmMain.printsalesdata.Recordset.Fields("time") = salesdata.Recordset.Fields("time")
        frmMain.printsalesdata.Recordset.Fields("itempurchased") = salesdata.Recordset.Fields("itempurchased")
        frmMain.printsalesdata.Recordset.Fields("cash") = salesdata.Recordset.Fields("cash")
        frmMain.printsalesdata.Recordset.Fields("change") = salesdata.Recordset.Fields("change")
        frmMain.printsalesdata.Recordset.Fields("total") = salesdata.Recordset.Fields("total")
        frmMain.printsalesdata.Recordset.Update
        frmMain.salesdata.Recordset.MoveNext
        Loop Until frmMain.salesdata.Recordset.EOF
        
        cmdSort.BackColor = &H8080FF
        cmdSort.Caption = "CANCEL"
    Else
        totalsales = 0
        
        MsgBox "No Record Exists In This Date!", vbCritical, ""
        salesdata.RecordSource = "select * from tblsales"
        salesdata.Refresh
        
        If Not salesdata.Recordset.EOF Then
            Do
            totalsales = totalsales + FormatNumber(frmMain.salesdata.Recordset.Fields("total"))
            frmMain.salesdata.Recordset.MoveNext
            Loop Until frmMain.salesdata.Recordset.EOF
        End If
    End If
    salesdata.Refresh
    lbltotalsales.Caption = FormatNumber(totalsales)
Else
    totalsales = 0

    salesdata.RecordSource = "select * from tblsales"
    salesdata.Refresh
     
    Do
    totalsales = totalsales + FormatNumber(frmMain.salesdata.Recordset.Fields("total"))
    frmMain.salesdata.Recordset.MoveNext
    Loop Until frmMain.salesdata.Recordset.EOF
    
    printsalesdata.RecordSource = "select * from tblprintsales"
    printsalesdata.Refresh
    
    Do
    frmMain.printsalesdata.Recordset.Delete
    printsalesdata.Refresh
    Loop Until printsalesdata.Recordset.EOF
    
    lbltotalsales.Caption = FormatNumber(totalsales)
    
    cmdSort.BackColor = &H82D1B0
    cmdSort.Caption = "SORT"
    salesdata.Refresh
End If
txtSearchOR.SetFocus
End Sub

Private Sub cmdStocks_Click()

frmProducts.Visible = False
frmStocks.Visible = True
cmdProducts.BackColor = &HC0C0C0
cmdStocks.BackColor = &HEAF1F4
txtsearchproduct.SetFocus
txtsearchproduct.Text = ""
stocksdata.RecordSource = "select * from tblstocks"
stocksdata.Refresh

End Sub

Private Sub cmdTransaction_Click() 'kailangan

printsalesdata.RecordSource = "select * from tblprintsales"
printsalesdata.Refresh

If Not printsalesdata.Recordset.EOF Then
    Do
    frmMain.printsalesdata.Recordset.Delete
    printsalesdata.Refresh
    Loop Until printsalesdata.Recordset.EOF
End If

cmdSort.Caption = "SORT"
cmdSort.BackColor = &H82D1B0

cmdTransaction.BackColor = &HDBF2E9
cmdSales.BackColor = &H82D1B0
cmdInventory.BackColor = &H82D1B0
cmdManageAccounts.BackColor = &H82D1B0
cmdLogOut.BackColor = &H82D1B0
frmTransaction.Visible = True
frmSales.Visible = False
frmInventory.Visible = False
frmManageAccounts.Visible = False

stocksdata.RecordSource = "select * from tblstocks"
stocksdata.Refresh

If frmCompleted.Visible = True Then
    cmdNewTransaction.SetFocus
Else
    txtsearchproduct_3.SetFocus
    txtsearchproduct_3.Text = ""
End If
End Sub

Private Sub cmdHome_Click() 'kailangan

'menubutton = cmdHome
cmdHome.BackColor = &HDBF2E9
cmdTransaction.BackColor = &H82D1B0
Command3.BackColor = &H82D1B0
cmdInventory.BackColor = &H82D1B0
cmdManageAccounts.BackColor = &H82D1B0
cmdLogOut.BackColor = &H82D1B0
frmHome.Visible = True
frmInventory.Visible = False
frmTransaction.Visible = False
frmManageAccounts.Visible = False

End Sub


Private Sub cmdUpdateAdmin_Click()
cmdUpdateAdmin.BackColor = &HDBF2E9

If accountsdata.Recordset.EOF Then
    MsgBox "Select Account To Update!", vbCritical, ""
    txtsearchaccount.Text = ""
    txtsearchaccount.SetFocus
    cmdUpdateAdmin.BackColor = &H82D1B0
    cmbUserType = "GENERAL"
    cmbAccountStatus = "GENERAL"
Else
    frmAskAdminPassword.whatbutton = "updateaccount"
    frmAskAdminPassword.Show
    frmMain.Enabled = False
    frmAskAdminPassword.txtAskAdminPass.SetFocus
End If
End Sub


Private Sub cmdUpdateProduct_Click()
If productsdatainv.Recordset.EOF Then
    MsgBox "Select Product To Update!", vbCritical, ""
    txtsearchproduct_2.Text = ""
    txtsearchproduct_2.SetFocus
Else
    frmAskAdminPassword.whatbutton = "updateproduct"
    frmAskAdminPassword.Show
    frmMain.Enabled = False
    frmAskAdminPassword.txtAskAdminPass.SetFocus
End If
End Sub

Private Sub cmdView_Click()
If accountsdata.Recordset.EOF Then
    MsgBox "Select Account To Update!", vbCritical, ""
    txtsearchaccount.Text = ""
    txtsearchaccount.SetFocus
    cmdView.BackColor = &H82D1B0
    cmbUserType = "GENERAL"
    cmbAccountStatus = "GENERAL"
Else
    cmdView.BackColor = &HDBF2E9
    frmAskAdminPassword.Show
    Me.Enabled = False
    frmAskAdminPassword.whatbutton = "viewaccount"
End If
End Sub

Private Sub cmdAddProduct_Click()
Me.Enabled = False
frmAskAdminPassword.Show
frmAskAdminPassword.whatbutton = "addproduct"
End Sub

Private Sub cmdViewProduct_Click()
If productsdatainv.Recordset.EOF Then
    MsgBox "Select Product To View!", vbCritical, ""
    txtsearchproduct_2.Text = ""
    txtsearchproduct_2.SetFocus
Else
    frmView.Show
    frmView.frmViewProduct.Visible = True
    frmMain.Enabled = False
    
    'migrate data
    frmView.lblprodid3.Caption = productsdatainv.Recordset.Fields("productid")
    frmView.lblbarcode3.Caption = productsdatainv.Recordset.Fields("barcode")
    frmView.lblprodname3.Caption = productsdatainv.Recordset.Fields("productname")
    frmView.lblprice3.Caption = productsdatainv.Recordset.Fields("price")
    frmView.txtdescription3.Text = productsdatainv.Recordset.Fields("description")
    frmView.txtdescription3.Enabled = False
End If
End Sub


Private Sub cmdViewRecord_Click()
If salesdata.Recordset.EOF Then
    MsgBox "Select Record To View!", vbCritical, ""
    txtSearchOR.Text = ""
    txtSearchOR.SetFocus
Else
    frmSalesRecord.productsalesdata.RecordSource = "select * from tblproductsales where orno= '" + salesdata.Recordset.Fields("orno") + "'"
    frmSalesRecord.productsalesdata.Refresh
    
    frmSalesRecord.printproductsalesdata.RecordSource = "select * from tblprintproductsales"
    frmSalesRecord.printproductsalesdata.Refresh
    
    If Not frmSalesRecord.productsalesdata.Recordset.EOF Then
        Do
        frmSalesRecord.printproductsalesdata.Recordset.Delete
        frmSalesRecord.printproductsalesdata.Refresh
        Loop Until frmSalesRecord.printproductsalesdata.Recordset.EOF
    End If
    
    frmSalesRecord.productsalesdata.RecordSource = "select * from tblproductsales where orno= '" + salesdata.Recordset.Fields("orno") + "'"
    frmSalesRecord.productsalesdata.Refresh
    
    Do
    With frmSalesRecord.printproductsalesdata.Recordset
        .AddNew
        .Fields("productid") = frmSalesRecord.productsalesdata.Recordset.Fields("productid")
        .Fields("orno") = frmSalesRecord.productsalesdata.Recordset.Fields("orno")
        .Fields("barcode") = frmSalesRecord.productsalesdata.Recordset.Fields("barcode")
        .Fields("productname") = frmSalesRecord.productsalesdata.Recordset.Fields("productname")
        .Fields("quantity") = frmSalesRecord.productsalesdata.Recordset.Fields("quantity")
        .Fields("price") = frmSalesRecord.productsalesdata.Recordset.Fields("price")
        .Fields("quantityprice") = frmSalesRecord.productsalesdata.Recordset.Fields("quantityprice")
        .Update
    End With
    frmSalesRecord.productsalesdata.Recordset.MoveNext
    Loop Until frmSalesRecord.productsalesdata.Recordset.EOF
    
    frmSalesRecord.txtSearchItem.Text = ""
    frmMain.Enabled = False
    frmSalesRecord.Show
    frmSalesRecord.txtSearchItem.SetFocus
End If
End Sub

Private Sub Command1_Click()
If stocksdata.Recordset.EOF Then
    MsgBox "Select Stock To Update Its Quantity!", vbCritical, ""
    txtsearchproduct.Text = ""
    txtsearchproduct.SetFocus
Else
    frmAskAdminPassword.whatbutton = "updatestockqty"
    frmAskAdminPassword.Show
    Me.Enabled = False
    frmAskAdminPassword.txtAskAdminPass.SetFocus
End If
End Sub

Private Sub cmdSales_Click()

totalsales = 0

salesdata.RecordSource = "select * from tblsales"
salesdata.Refresh


If Not salesdata.Recordset.EOF Then
    Do
    totalsales = totalsales + FormatNumber(frmMain.salesdata.Recordset.Fields("total"))
    frmMain.salesdata.Recordset.MoveNext
    Loop Until frmMain.salesdata.Recordset.EOF
End If

printsalesdata.RecordSource = "select * from tblprintsales"
printsalesdata.Refresh

If Not printsalesdata.Recordset.EOF Then
    Do
    frmMain.printsalesdata.Recordset.Delete
    printsalesdata.Refresh
    Loop Until printsalesdata.Recordset.EOF
End If

cmdSort.Caption = "SORT"
cmdSort.BackColor = &H82D1B0
    
lbltotalsales.Caption = FormatNumber(totalsales)

salesdata.RecordSource = "select * from tblsales"
salesdata.Refresh

txtSearchOR.Text = ""

cmdSales.BackColor = &HDBF2E9
cmdTransaction.BackColor = &H82D1B0
cmdInventory.BackColor = &H82D1B0
cmdManageAccounts.BackColor = &H82D1B0
cmdLogOut.BackColor = &H82D1B0
frmTransaction.Visible = False
frmInventory.Visible = False
frmManageAccounts.Visible = False
frmSales.Visible = True

txtSearchOR.SetFocus

'MULA DITO - PARA SA TRANSACTION REFRESH
'PARA SA TRANSACTION GUI
    frmMain.frmCompleted.Visible = False
    frmMain.Frame1.Enabled = True
    frmMain.frmCompleted.Enabled = False
    frmMain.lbltotal.FontSize = 24
    frmPopUps.total = 0
    
'COLORS
    frmMain.cmdAddToCart.BackColor = &H82D1B0
    frmMain.cmdEditQty.BackColor = &H82D1B0
    frmMain.cmdRemoveItem.BackColor = &HC0C0FF
    frmMain.cmdSearchCart.BackColor = &H82D1B0
    frmMain.cmdCancelTransaction.BackColor = &HC0C0FF
    frmMain.cmdPay.BackColor = &H80FF80
    frmMain.DBGrid3.BackColor = &H80000005
    frmMain.DBGrid4.BackColor = &H80000005
    frmMain.Frame6.BackColor = &H3E472C
    frmMain.Frame5.BackColor = &H3E472C
    frmMain.Frame7.BackColor = &H3E472C
    frmMain.frmTransaction.BackColor = &H82D1B0
    frmMain.Image3.Picture = LoadPicture(App.Path & "/images/cart_logo_2.jpg")
    
'NEW TRANSACTION AUTO
    transactdata.RecordSource = "select * from tbltransact"
    transactdata.Refresh
    
    If Not transactdata.Recordset.EOF Then
        Do
        transactdata.Recordset.Delete
        transactdata.Refresh
        Loop Until transactdata.Recordset.EOF
    End If

'DATA REFRESHES
    stocksdata.RecordSource = "select * from tblstocks"
    stocksdata.Refresh
    transactdata.RecordSource = "select * from tbltransact"
    transactdata.Refresh
    
'LABEL REFRESHES
    lbltotal.Caption = "0.00"
    lblchange.Caption = "0.00"
'HANGGANG DITO LANG
End Sub




Private Sub dtsales_Change()
txtSearchOR.Text = ""

printsalesdata.RecordSource = "select * from tblprintsales"
printsalesdata.Refresh

If Not printsalesdata.Recordset.EOF Then
    Do
    frmMain.printsalesdata.Recordset.Delete
    printsalesdata.Refresh
    Loop Until printsalesdata.Recordset.EOF
End If

cmdSort.BackColor = &H82D1B0
cmdSort.Caption = "SORT"
End Sub

Private Sub Form_Load() 'kailangan

frmInventory.Visible = False
frmTransaction.Visible = True
frmManageAccounts.Visible = False

cmbUserType = "GENERAL"
cmbAccountStatus = "GENERAL"
cmbCategories = "GENERAL"

End Sub

Private Sub Timer1_Timer() 'kailangan

'date and time should be here
Timer1.Enabled = True
lblTime.Caption = Time
lblDate.Caption = Format(Date, "mmmm dd, yyyy") 'date format
lblDay.Caption = Format(Date, "dddd")

End Sub

Private Sub txtSearch_p_KeyPress(KeyAscii As Integer) 'di pa sure

If KeyAscii = 13 Then
    cmdSearch_p.SetFocus
End If
End Sub

Private Sub txtsearchaccount_Change()

If txtsearchaccount.Text = Empty Then
    accountsdata.RecordSource = "select * from tblaccounts"
    accountsdata.Refresh
End If

If cmbAccountStatus.Text = "GENERAL" And cmbUserType.Text = "GENERAL" Then
    If Not txtsearchaccount.Text = Empty Then
        accountsdata.RecordSource = "select * from tblaccounts where accid= '" + txtsearchaccount.Text + "'"
        accountsdata.Refresh
    Else
        accountsdata.RecordSource = "select * from tblaccounts"
        accountsdata.Refresh
    End If
Else
    If cmbAccountStatus.Text = "ACTIVE" And cmbUserType.Text = "GENERAL" Then
        If Not txtsearchaccount.Text = Empty Then
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and accid= '" + txtsearchaccount + "'"
            accountsdata.Refresh
        Else
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "'"
            accountsdata.Refresh
        End If
    Else
        If cmbAccountStatus.Text = "NOT ACTIVE" And cmbUserType.Text = "GENERAL" Then
            If Not txtsearchaccount.Text = Empty Then
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and accid= '" + txtsearchaccount.Text + "'"
                accountsdata.Refresh
            Else
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "'"
                accountsdata.Refresh
            End If
        End If
    End If
End If


If cmbAccountStatus.Text = "GENERAL" And cmbUserType.Text = "ADMIN" Then
    If Not txtsearchaccount.Text = Empty Then
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount + "'"
        accountsdata.Refresh
    Else
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "'"
        accountsdata.Refresh
    End If
Else
    If cmbAccountStatus.Text = "ACTIVE" And cmbUserType.Text = "ADMIN" Then
        If Not txtsearchaccount.Text = Empty Then
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
            accountsdata.Refresh
        Else
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
            accountsdata.Refresh
        End If
    Else
        If cmbAccountStatus.Text = "NOT ACTIVE" And cmbUserType.Text = "ADMIN" Then
            If Not txtsearchaccount.Text = Empty Then
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
                accountsdata.Refresh
            Else
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
                accountsdata.Refresh
            End If
        End If
    End If
End If


If cmbAccountStatus.Text = "GENERAL" And cmbUserType.Text = "USER" Then
    If Not txtsearchaccount.Text = Empty Then
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
        accountsdata.Refresh
    Else
        accountsdata.RecordSource = "select * from tblaccounts where usertype= '" + cmbUserType + "'"
        accountsdata.Refresh
    End If
Else
    If cmbAccountStatus.Text = "ACTIVE" And cmbUserType.Text = "USER" Then
        If Not txtsearchaccount.Text = Empty Then
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
            accountsdata.Refresh
        Else
            accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
            accountsdata.Refresh
        End If
    Else
        If cmbAccountStatus.Text = "NOT ACTIVE" And cmbUserType.Text = "USER" Then
            If Not txtsearchaccount.Text = Empty Then
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "' and accid= '" + txtsearchaccount.Text + "'"
                accountsdata.Refresh
            Else
                accountsdata.RecordSource = "select * from tblaccounts where status= '" + cmbAccountStatus + "' and usertype= '" + cmbUserType + "'"
                accountsdata.Refresh
            End If
        End If
    End If
End If
End Sub

Private Sub txtsearchaccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789._ ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtSearchOR_Change()
If txtSearchOR.Text = Empty Then
    If cmdSort.Caption = "CANCEL" Then
        salesdata.RecordSource = "select * from tblsales where date= '" + sortdate + "'"
        salesdata.Refresh
    Else
        salesdata.RecordSource = "select * from tblsales"
        salesdata.Refresh
    End If
Else
    If cmdSort.Caption = "CANCEL" Then
        salesdata.RecordSource = "select * from tblsales where orno= '" + txtSearchOR.Text + "' and date= '" + sortdate + "'"
        salesdata.Refresh
    Else
        salesdata.RecordSource = "select * from tblsales where orno= '" + txtSearchOR.Text + "'"
        salesdata.Refresh
    End If
End If
End Sub

Private Sub txtSearchOR_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtsearchproduct_2_Change()

If txtsearchproduct_2.Text = Empty Then
    productsdatainv.RecordSource = "select * from tblproducts"
    productsdatainv.Refresh
Else
    productsdatainv.RecordSource = "select * from tblproducts where productid= '" + txtsearchproduct_2.Text + "'"
    productsdatainv.Refresh
    
    If productsdatainv.Recordset.EOF Then
        productsdatainv.RecordSource = "select * from tblproducts where productname= '" + txtsearchproduct_2.Text + "'"
        productsdatainv.Refresh
        
        If productsdatainv.Recordset.EOF Then
            productsdatainv.RecordSource = "select * from tblproducts where barcode= '" + txtsearchproduct_2.Text + "'"
            productsdatainv.Refresh
        End If
    End If
End If
End Sub

Private Sub txtsearchproduct_2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789()-_,.#!&%$* ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtsearchproduct_3_Change()
If txtsearchproduct_3.Text = Empty Then
    stocksdata.RecordSource = "select * from tblstocks"
    stocksdata.Refresh
Else
    stocksdata.RecordSource = "select * from tblstocks where productid= '" + txtsearchproduct_3.Text + "'"
    stocksdata.Refresh
    
    If stocksdata.Recordset.EOF Then
        stocksdata.RecordSource = "select * from tblstocks where productname= '" + txtsearchproduct_3.Text + "'"
        stocksdata.Refresh
        
        If stocksdata.Recordset.EOF Then
            stocksdata.RecordSource = "select * from tblstocks where barcode= '" + txtsearchproduct_3.Text + "'"
            stocksdata.Refresh
        End If
    End If
End If
End Sub

Private Sub txtsearchproduct_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAddToCart.SetFocus
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789()-_,.#!&%$* ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtsearchproduct_Change()

If txtsearchproduct.Text = Empty Then
    stocksdata.RecordSource = "select * from tblstocks"
    stocksdata.Refresh
Else
    stocksdata.RecordSource = "select * from tblstocks where productid= '" + txtsearchproduct.Text + "'"
    stocksdata.Refresh
    
    If stocksdata.Recordset.EOF Then
        stocksdata.RecordSource = "select * from tblstocks where productname= '" + txtsearchproduct.Text + "'"
        stocksdata.Refresh
        
        If stocksdata.Recordset.EOF Then
            stocksdata.RecordSource = "select * from tblstocks where barcode= '" + txtsearchproduct.Text + "'"
            stocksdata.Refresh
        End If
    End If
End If
End Sub

Private Sub txtsearchproduct_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789()-_,.#!&%$* ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
