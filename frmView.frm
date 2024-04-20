VERSION 5.00
Begin VB.Form frmView 
   BackColor       =   &H003E472C&
   BorderStyle     =   0  'None
   ClientHeight    =   8175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmUpdateProduct 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton Command1 
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   375
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H0082D1B0&
         Height          =   7095
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtprodname2 
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
            MaxLength       =   50
            TabIndex        =   13
            Top             =   2400
            Width           =   5295
         End
         Begin VB.TextBox txtprice2 
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
            MaxLength       =   8
            TabIndex        =   14
            Top             =   3360
            Width           =   2415
         End
         Begin VB.TextBox txtdescription2 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   720
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   4200
            Width           =   5295
         End
         Begin VB.CommandButton cmdSaveUpdate 
            BackColor       =   &H00DBF2E9&
            Caption         =   "SAVE CHANGES"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   5760
            Width           =   1695
         End
         Begin VB.TextBox txtbarcode2 
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
            MaxLength       =   50
            TabIndex        =   12
            Top             =   1560
            Width           =   5295
         End
         Begin VB.Label lblmessage4_3 
            AutoSize        =   -1  'True
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
            Left            =   2520
            TabIndex        =   70
            Top             =   1320
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblprodid2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1920
            TabIndex        =   58
            Top             =   840
            Width           =   105
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT ID:"
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
            TabIndex        =   57
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT NAME:"
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
            TabIndex        =   56
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRICE:"
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
            TabIndex        =   55
            Top             =   3120
            Width           =   510
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION:"
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
            TabIndex        =   54
            Top             =   3960
            Width           =   1125
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT BARCODE:"
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
            TabIndex        =   53
            Top             =   1320
            Width           =   1635
         End
         Begin VB.Label lblmessage4_1 
            AutoSize        =   -1  'True
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
            Left            =   2160
            TabIndex        =   52
            Top             =   2160
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage4_2 
            AutoSize        =   -1  'True
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
            Left            =   1320
            TabIndex        =   51
            Top             =   3120
            Visible         =   0   'False
            Width           =   45
         End
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   2160
         Picture         =   "frmView.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UPDATE PRODUCT"
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
         Left            =   2760
         TabIndex        =   59
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frmAddProduct 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdCancel2 
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0082D1B0&
         Height          =   7095
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   6855
         Begin VB.TextBox txtprice1 
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
            MaxLength       =   8
            TabIndex        =   3
            Top             =   3360
            Width           =   2415
         End
         Begin VB.TextBox txtBarcode1 
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
            MaxLength       =   50
            TabIndex        =   1
            Top             =   1560
            Width           =   5295
         End
         Begin VB.Data productsdatainv 
            Caption         =   "PRODUCTS"
            Connect         =   "Access"
            DatabaseName    =   "database\database_sabana.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   4320
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "tblproducts"
            Top             =   240
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CommandButton cmdadd 
            BackColor       =   &H00DBF2E9&
            Caption         =   "ADD PRODUCT"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   5760
            Width           =   1695
         End
         Begin VB.TextBox txtdescription1 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   720
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   4200
            Width           =   5295
         End
         Begin VB.TextBox txtprodname1 
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
            MaxLength       =   50
            TabIndex        =   2
            Top             =   2400
            Width           =   5295
         End
         Begin VB.Label lblmessage3_3 
            AutoSize        =   -1  'True
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
            Left            =   2520
            TabIndex        =   69
            Top             =   1320
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage3_2 
            AutoSize        =   -1  'True
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
            Left            =   1320
            TabIndex        =   48
            Top             =   3120
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage3_1 
            AutoSize        =   -1  'True
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
            Left            =   2160
            TabIndex        =   47
            Top             =   2160
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT BARCODE:"
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
            TabIndex        =   46
            Top             =   1320
            Width           =   1635
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION:"
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
            TabIndex        =   45
            Top             =   3960
            Width           =   1125
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRICE:"
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
            TabIndex        =   44
            Top             =   3120
            Width           =   510
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT NAME:"
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
            TabIndex        =   43
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT ID:"
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
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label lblprodid1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1920
            TabIndex        =   41
            Top             =   840
            Width           =   105
         End
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   2400
         Picture         =   "frmView.frx":264B6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADD PRODUCT"
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
         Left            =   3030
         TabIndex        =   39
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.Frame frmViewAccount 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdCancel1 
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   375
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0082D1B0&
         Caption         =   "ACCOUNT INFORMATION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   6615
         Begin VB.Frame Frame11 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   2400
            TabIndex        =   107
            Top             =   4560
            Width           =   3855
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "EMAIL ADDRESS:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   109
               Top             =   0
               Width           =   1350
            End
            Begin VB.Label lblemail1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               TabIndex        =   108
               Top             =   240
               Width           =   105
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   360
            TabIndex        =   104
            Top             =   3000
            Width           =   5895
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "LAST NAME:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   106
               Top             =   0
               Width           =   1005
            End
            Begin VB.Label lbllast1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               TabIndex        =   105
               Top             =   240
               Width           =   105
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   360
            TabIndex        =   101
            Top             =   2400
            Width           =   5895
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MIDDLE NAME:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   103
               Top             =   0
               Width           =   1245
            End
            Begin VB.Label lblmiddle1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               TabIndex        =   102
               Top             =   240
               Width           =   105
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   360
            TabIndex        =   98
            Top             =   1800
            Width           =   5895
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "FIRST NAME:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   100
               Top             =   0
               Width           =   1035
            End
            Begin VB.Label lblfirst1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               TabIndex        =   99
               Top             =   240
               Width           =   105
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   360
            TabIndex        =   95
            Top             =   1080
            Width           =   5895
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ACCOUNT USERNAME:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   120
               TabIndex        =   97
               Top             =   0
               Width           =   1830
            End
            Begin VB.Label lblusername1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               TabIndex        =   96
               Top             =   240
               Width           =   105
            End
         End
         Begin VB.CommandButton cmdsave2 
            BackColor       =   &H00DBF2E9&
            Caption         =   "CLOSE"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   5640
            Width           =   1695
         End
         Begin VB.Label lblaccountid1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   600
            TabIndex        =   37
            Top             =   720
            Width           =   105
         End
         Begin VB.Label lblstatus1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   225
            Left            =   5400
            TabIndex        =   36
            Top             =   480
            Width           =   105
         End
         Begin VB.Label lblsex1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   5040
            TabIndex        =   35
            Top             =   4080
            Width           =   105
         End
         Begin VB.Label lblbday1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   2640
            TabIndex        =   34
            Top             =   4080
            Width           =   105
         End
         Begin VB.Label lblphone1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   600
            TabIndex        =   33
            Top             =   4800
            Width           =   105
         End
         Begin VB.Label lblage1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   600
            TabIndex        =   32
            Top             =   4080
            Width           =   105
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STATUS:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4560
            TabIndex        =   31
            Top             =   480
            Width           =   660
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PHONE:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   30
            Top             =   4560
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SEX:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4920
            TabIndex        =   29
            Top             =   3840
            Width           =   345
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BIRTHDAY:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2520
            TabIndex        =   28
            Top             =   3840
            Width           =   870
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AGE:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   27
            Top             =   3840
            Width           =   390
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT ID:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   480
            TabIndex        =   25
            Top             =   480
            Width           =   1080
         End
      End
      Begin VB.Label lblusertype1 
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
         Left            =   3240
         TabIndex        =   26
         Top             =   240
         Width           =   1230
      End
      Begin VB.Image imgusertype1 
         Height          =   720
         Left            =   2400
         Picture         =   "frmView.frx":417C8
         Stretch         =   -1  'True
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.Frame frmUpdateStock 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   85
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdCancel3 
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H0082D1B0&
         Height          =   7095
         Left            =   120
         TabIndex        =   86
         Top             =   720
         Width           =   6855
         Begin VB.Frame Frame13 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   600
            TabIndex        =   113
            Top             =   1440
            Width           =   5655
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PRODUCT BARCODE:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   0
               TabIndex        =   115
               Top             =   0
               Width           =   1890
            End
            Begin VB.Label lblbarcode5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               TabIndex        =   114
               Top             =   360
               Width           =   105
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   600
            TabIndex        =   110
            Top             =   2280
            Width           =   5655
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PRODUCT NAME:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   0
               TabIndex        =   112
               Top             =   0
               Width           =   1590
            End
            Begin VB.Label lblprodname5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               TabIndex        =   111
               Top             =   360
               Width           =   105
            End
         End
         Begin VB.CommandButton cmdSaveUpdateStock 
            BackColor       =   &H00DBF2E9&
            Caption         =   "SAVE CHANGES"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   5880
            Width           =   1695
         End
         Begin VB.TextBox txtdescription5 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   600
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   75
            Top             =   4320
            Width           =   5655
         End
         Begin VB.TextBox txtstocks5 
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
            Left            =   3720
            MaxLength       =   6
            TabIndex        =   7
            Top             =   3480
            Width           =   2535
         End
         Begin VB.Label lblprice5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   840
            TabIndex        =   93
            Top             =   3480
            Width           =   105
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   92
            Top             =   3960
            Width           =   1290
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRICE:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   91
            Top             =   3120
            Width           =   585
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT ID:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   90
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lblprodid5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   840
            TabIndex        =   89
            Top             =   960
            Width           =   105
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STOCKS:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3480
            TabIndex        =   88
            Top             =   3120
            Width           =   765
         End
         Begin VB.Label lblmessage7_1 
            AutoSize        =   -1  'True
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
            Left            =   4320
            TabIndex        =   87
            Top             =   3135
            Visible         =   0   'False
            Width           =   45
         End
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   1680
         Picture         =   "frmView.frx":4D87A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UPDATE STOCK QUANTITY"
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
         Left            =   2280
         TabIndex        =   94
         Top             =   240
         Width           =   3090
      End
   End
   Begin VB.Frame frmAddNewStock 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   71
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Frame Frame5 
         BackColor       =   &H0082D1B0&
         Height          =   7095
         Left            =   120
         TabIndex        =   72
         Top             =   720
         Width           =   6855
         Begin VB.Frame Frame15 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   600
            TabIndex        =   119
            Top             =   2640
            Width           =   5655
            Begin VB.Label lblprodname4 
               AutoSize        =   -1  'True
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
               Height          =   225
               Left            =   240
               TabIndex        =   121
               Top             =   360
               Width           =   45
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PRODUCT NAME:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   0
               TabIndex        =   120
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   600
            TabIndex        =   116
            Top             =   1800
            Width           =   5655
            Begin VB.Label lblbarcode4 
               AutoSize        =   -1  'True
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
               Height          =   225
               Left            =   240
               TabIndex        =   118
               Top             =   360
               Width           =   45
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PRODUCT BARCODE:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   0
               TabIndex        =   117
               Top             =   0
               Width           =   1890
            End
         End
         Begin VB.ComboBox cmbProducts 
            BackColor       =   &H00FFFFFF&
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
            ItemData        =   "frmView.frx":73D30
            Left            =   3720
            List            =   "frmView.frx":73D32
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtstocks4 
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
            Left            =   3720
            MaxLength       =   6
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   3840
            Width           =   2535
         End
         Begin VB.TextBox txtdescription4 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   600
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   4680
            Width           =   5655
         End
         Begin VB.CommandButton cmdAddNewStocks 
            BackColor       =   &H00DBF2E9&
            Caption         =   "ADD STOCK"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   6120
            Width           =   1695
         End
         Begin VB.Label lblmessage5_1 
            AutoSize        =   -1  'True
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
            Left            =   5280
            TabIndex        =   84
            Top             =   975
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage5_2 
            AutoSize        =   -1  'True
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
            Left            =   4320
            TabIndex        =   83
            Top             =   3495
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SELECT PRODUCT:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3480
            TabIndex        =   82
            Top             =   960
            Width           =   1635
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "STOCKS:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3480
            TabIndex        =   81
            Top             =   3480
            Width           =   765
         End
         Begin VB.Label lblprodid4 
            AutoSize        =   -1  'True
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
            Height          =   225
            Left            =   840
            TabIndex        =   79
            Top             =   1320
            Width           =   45
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT ID:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   78
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRICE:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   77
            Top             =   3480
            Width           =   585
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   76
            Top             =   4320
            Width           =   1290
         End
         Begin VB.Label lblprice4 
            AutoSize        =   -1  'True
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
            Height          =   225
            Left            =   840
            TabIndex        =   74
            Top             =   3840
            Width           =   45
         End
      End
      Begin VB.CommandButton cmdExit4 
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image4 
         Height          =   375
         Left            =   2280
         Picture         =   "frmView.frx":73D34
         Stretch         =   -1  'True
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADD NEW STOCK"
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
         Left            =   2880
         TabIndex        =   80
         Top             =   240
         Width           =   2010
      End
   End
   Begin VB.Frame frmViewProduct 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   60
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CommandButton cmdExit3 
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Width           =   375
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0082D1B0&
         Height          =   7095
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   6855
         Begin VB.Frame Frame17 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   600
            TabIndex        =   125
            Top             =   2280
            Width           =   5655
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PRODUCT NAME:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   0
               TabIndex        =   127
               Top             =   0
               Width           =   1590
            End
            Begin VB.Label lblprodname3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               TabIndex        =   126
               Top             =   360
               Width           =   105
            End
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   600
            TabIndex        =   122
            Top             =   1440
            Width           =   5655
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PRODUCT BARCODE:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   0
               TabIndex        =   124
               Top             =   0
               Width           =   1890
            End
            Begin VB.Label lblbarcode3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               TabIndex        =   123
               Top             =   360
               Width           =   105
            End
         End
         Begin VB.CommandButton cmdClose 
            BackColor       =   &H00DBF2E9&
            Caption         =   "CLOSE"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   5880
            Width           =   1695
         End
         Begin VB.TextBox txtdescription3 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   600
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   62
            Top             =   4320
            Width           =   5655
         End
         Begin VB.Label lblprice3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   840
            TabIndex        =   68
            Top             =   3480
            Width           =   105
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPTION:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   66
            Top             =   3960
            Width           =   1290
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRICE:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   65
            Top             =   3120
            Width           =   585
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCT ID:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   600
            TabIndex        =   64
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lblprodid3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   840
            TabIndex        =   63
            Top             =   960
            Width           =   105
         End
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW PRODUCT"
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
         Left            =   2520
         TabIndex        =   67
         Top             =   240
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public prodid As Integer

Private Sub cmbProducts_Click()

productsdatainv.RecordSource = "select * from tblproducts where barcode= '" + cmbProducts + "'"
productsdatainv.Refresh

    If cmbProducts.Text = productsdatainv.Recordset.Fields("barcode") Then
        lblprodid4.Caption = productsdatainv.Recordset.Fields("productid")
        lblbarcode4.Caption = productsdatainv.Recordset.Fields("barcode")
        lblprodname4.Caption = productsdatainv.Recordset.Fields("productname")
        lblprice4.Caption = productsdatainv.Recordset.Fields("price")
        txtdescription4.Text = productsdatainv.Recordset.Fields("description")
    End If
lblmessage5_1.Visible = False
End Sub

Private Sub cmdadd_Click()

If Len(txtBarcode1.Text) = Empty Then
    lblmessage3_3.Caption = "(ENTER BARCODE)"
    lblmessage3_3.Visible = True
    txtBarcode1.SetFocus
Else
    If Len(txtprodname1.Text) <= 1 Then
        lblmessage3_1.Caption = "(ENTER VALID PRODUCT NAME)"
        lblmessage3_1.Visible = True
        txtprodname1.SetFocus
    Else
        If txtprice1.Text = Empty Then
            lblmessage3_2.Caption = "(ENTER PRICE)"
            lblmessage3_2.Visible = True
            txtprice1.SetFocus
        Else
        
            If Val(txtprice1.Text) = 0 Then
                lblmessage3_2.Caption = "(ENTER VALID PRICE)"
                lblmessage3_2.Visible = True
                txtprice1.SetFocus
            Else
                productsdatainv.RecordSource = "select * from tblproducts where barcode= '" + txtBarcode1.Text + "'"
                productsdatainv.Refresh
                
                If productsdatainv.Recordset.EOF Then
                    MsgBox "Product Added!", vbInformation, ""
                    
                    With frmMain.productsdatainv.Recordset
                        .AddNew
                        .Fields("productid") = lblprodid1.Caption
                        .Fields("barcode") = txtBarcode1.Text
                        .Fields("productname") = txtprodname1.Text
                        .Fields("price") = Format(Val(txtprice1.Text), "0.00")
                        .Fields("description") = txtdescription1.Text
                        .Update
                    End With
                    
                    productsdatainv.RecordSource = "select * from tblproducts"
                    productsdatainv.Refresh
                    
                    Do
                    prodid = productsdatainv.Recordset.Fields("productid")
                    productsdatainv.Recordset.MoveNext
                    Loop Until productsdatainv.Recordset.EOF
                            
                    prodid = prodid + 1
                    lblprodid1.Caption = prodid
                    
                    txtBarcode1.SetFocus
                    
                    'clear
                    txtBarcode1.Text = ""
                    txtprodname1.Text = ""
                    txtprice1.Text = ""
                    txtdescription1.Text = ""
                Else
                    lblmessage3_3.Caption = "(BARCODE IS ALREADY TAKEN)"
                    lblmessage3_3.Visible = True
                    txtBarcode1.SetFocus
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub cmdAddCategory_Click()
Me.Enabled = False
frmCategories.Show
End Sub

Private Sub cmdAddNewStocks_Click()
If cmbProducts.Text = "" Then
    lblmessage5_1.Caption = "(SELECT)"
    lblmessage5_1.Visible = True
    cmbProducts.SetFocus
Else
    If txtstocks4.Text = Empty Then
        lblmessage5_2.Caption = "(ENTER STOCKS)"
        lblmessage5_2.Visible = True
        txtstocks4.SetFocus
    Else
        If Val(txtstocks4.Text) = 0 Then
            lblmessage5_2.Caption = "(ENTER VALID STOCKS)"
            lblmessage5_2.Visible = True
            txtstocks4.SetFocus
        Else
            frmMain.stocksdata.RecordSource = "select * from tblstocks where productid= '" + lblprodid4.Caption + "'"
            frmMain.stocksdata.Refresh
            
            If frmMain.stocksdata.Recordset.EOF Then
                MsgBox "Stocks Added!", vbInformation, ""
                
                frmMain.stocksdata.Recordset.AddNew
                frmMain.stocksdata.Recordset.Fields("productid") = lblprodid4.Caption
                frmMain.stocksdata.Recordset.Fields("barcode") = lblbarcode4.Caption
                frmMain.stocksdata.Recordset.Fields("productname") = lblprodname4.Caption
                frmMain.stocksdata.Recordset.Fields("price") = lblprice4.Caption
                frmMain.stocksdata.Recordset.Fields("description") = txtdescription4.Text
                frmMain.stocksdata.Recordset.Fields("stocks") = Val(txtstocks4.Text)
                frmMain.stocksdata.Recordset.Update
                
                frmMain.stocksdata.RecordSource = "select * from tblstocks"
                frmMain.stocksdata.Refresh
                
                txtstocks4.Text = ""
                cmbProducts.SetFocus
            Else
                frmMain.stocksdata.RecordSource = "select * from tblstocks"
                frmMain.stocksdata.Refresh
                
                MsgBox "Stock Already Existing! If You Want To Change Stock Quantity, Click (Update Stock Quantity).", vbCritical + vbOKOnly, ""
                lblmessage5_1.Caption = "(SELECT)"
                lblmessage5_1.Visible = True
                cmbProducts.SetFocus
                
            End If
        End If
    End If
End If
End Sub

Private Sub cmdCancel1_Click()
frmMain.cmdView.BackColor = &H82D1B0
frmMain.Enabled = True
frmMain.cmbAccountStatus = "GENERAL"
frmMain.cmbUserType = "GENERAL"
frmMain.txtsearchaccount.Text = ""
frmMain.txtsearchaccount.SetFocus
frmMain.Show
Unload Me
End Sub

Private Sub cmdCancel1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmMain.cmdView.BackColor = &H82D1B0
    frmMain.Enabled = True
    frmMain.cmbAccountStatus = "GENERAL"
    frmMain.cmbUserType = "GENERAL"
    frmMain.txtsearchaccount.Text = ""
    frmMain.Show
    Unload Me
End If
End Sub

Private Sub cmdCancel2_Click()
Dim reply As String
reply = MsgBox("Close?", vbQuestion + vbYesNo, "")

If reply = vbYes Then
    frmMain.Enabled = True
    frmMain.productsdatainv.Refresh
    frmMain.txtsearchproduct_2.Text = ""
    frmMain.txtsearchproduct_2.SetFocus
    frmMain.Show
    Unload Me
End If
End Sub

Private Sub cmdCancel3_Click()
Dim reply6 As String

reply6 = MsgBox("Close?", vbQuestion + vbYesNo, "")

If reply6 = vbYes Then
    frmMain.Enabled = True
    frmMain.stocksdata.RecordSource = "select * from tblstocks"
    frmMain.stocksdata.Refresh
    frmMain.txtsearchproduct.Text = ""
    frmMain.txtsearchproduct.SetFocus
    frmMain.Show
    Unload Me
End If
End Sub

Private Sub cmdClose_Click()
frmMain.Enabled = True
frmMain.txtsearchproduct_2.Text = ""
frmMain.txtsearchproduct_2.SetFocus
frmMain.Show
Unload Me
End Sub

Private Sub cmdExit3_Click()
frmMain.Enabled = True
frmMain.txtsearchproduct_2.Text = ""
frmMain.txtsearchproduct_2.SetFocus
frmMain.Show
Unload Me
End Sub

Private Sub cmdExit4_Click()
Dim reply As String
reply = MsgBox("Close?", vbQuestion + vbYesNo, "")

If reply = vbYes Then
    frmMain.Enabled = True
    frmMain.stocksdata.RecordSource = "select * from tblstocks"
    frmMain.stocksdata.Refresh
    frmMain.txtsearchproduct.Text = ""
    frmMain.txtsearchproduct.SetFocus
    frmMain.Show
    Unload Me
End If
End Sub

Private Sub cmdsave2_Click()
frmMain.cmdView.BackColor = &H82D1B0
frmMain.Enabled = True
frmMain.cmbAccountStatus = "GENERAL"
frmMain.cmbUserType = "GENERAL"
frmMain.txtsearchaccount.Text = ""
frmMain.txtsearchaccount.SetFocus
frmMain.Show
Unload Me
End Sub

Private Sub cmdsave2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    frmMain.cmdView.BackColor = &H82D1B0
    frmMain.Enabled = True
    frmMain.cmbAccountStatus = "GENERAL"
    frmMain.cmbUserType = "GENERAL"
    frmMain.txtsearchaccount.Text = ""
    frmMain.Show
Unload Me
End If
End Sub

Private Sub cmdSaveUpdate_Click()
If Len(txtbarcode2.Text) = Empty Then
    lblmessage4_3.Caption = "(ENTER BARCODE)"
    lblmessage4_3.Visible = True
    txtbarcode2.SetFocus
Else
    If Len(txtprodname2.Text) <= 1 Then
        lblmessage4_1.Caption = "(ENTER VALID PRODUCT NAME)"
        lblmessage4_1.Visible = True
        txtprodname2.SetFocus
    Else
        If txtprice2.Text = Empty Then
            lblmessage4_2.Caption = "(ENTER PRICE)"
            lblmessage4_2.Visible = True
            txtprice2.SetFocus
        Else
            productsdatainv.RecordSource = "select * from tblproducts where barcode= '" + txtbarcode2.Text + "'"
            productsdatainv.Refresh
            
            If productsdatainv.Recordset.EOF Then
                If Val(txtprice2.Text) = 0 Then
                    lblmessage4_2.Caption = "(ENTER VALID PRICE)"
                    lblmessage4_2.Visible = True
                    txtprice2.SetFocus
                Else
                    MsgBox "Product Updated!", vbInformation, ""
                    
                    cmdSaveUpdate.Caption = "SAVED"
                    cmdSaveUpdate.Enabled = False
                    
                    With frmMain.productsdatainv.Recordset
                        .Edit
                        .Fields("barcode") = txtbarcode2.Text
                        .Fields("productname") = txtprodname2.Text
                        .Fields("price") = Format(Val(txtprice2.Text), "0.00")
                        .Fields("description") = txtdescription2.Text
                        .Update
                    End With
                    
                    frmMain.stocksdata.RecordSource = "select * from tblstocks where productid= '" + lblprodid2.Caption + "'"
                    frmMain.stocksdata.Refresh
                    
                    If Not frmMain.stocksdata.Recordset.EOF Then
                        With frmMain.stocksdata.Recordset
                            .Edit
                            .Fields("barcode") = txtbarcode2.Text
                            .Fields("productname") = txtprodname2.Text
                            .Fields("price") = Format(Val(txtprice2.Text), "0.00")
                            .Fields("description") = txtdescription2.Text
                            .Update
                        End With
                    End If
                End If
                frmMain.Show
                frmMain.Enabled = True
                Unload Me
            Else
                If txtbarcode2.Text = frmMain.productsdatainv.Recordset.Fields("barcode") Then
                    If Val(txtprice2.Text) = 0 Then
                        lblmessage4_2.Caption = "(ENTER VALID PRICE)"
                        lblmessage4_2.Visible = True
                        txtprice2.SetFocus
                    Else
                        MsgBox "Product Updated!", vbInformation, ""
                        
                        cmdSaveUpdate.Caption = "SAVED"
                        cmdSaveUpdate.Enabled = False
                        
                        With frmMain.productsdatainv.Recordset
                            .Edit
                            .Fields("barcode") = txtbarcode2.Text
                            .Fields("productname") = txtprodname2.Text
                            .Fields("price") = Format(Val(txtprice2.Text), "0.00")
                            .Fields("description") = txtdescription2.Text
                            .Update
                        End With
                        
                        frmMain.stocksdata.RecordSource = "select * from tblstocks where productid= '" + lblprodid2.Caption + "'"
                        frmMain.stocksdata.Refresh
                    
                        If Not frmMain.stocksdata.Recordset.EOF Then
                            With frmMain.stocksdata.Recordset
                                .Edit
                                .Fields("barcode") = txtbarcode2.Text
                                .Fields("productname") = txtprodname2.Text
                                .Fields("price") = Format(Val(txtprice2.Text), "0.00")
                                .Fields("description") = txtdescription2.Text
                                .Update
                            End With
                        End If
                        frmMain.Show
                        frmMain.Enabled = True
                        Unload Me
                    End If
                Else
                    lblmessage4_3.Caption = "(BARCODE IS ALREADY IN USE)"
                    lblmessage4_3.Visible = True
                    txtbarcode2.SetFocus
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub cmdSaveUpdateStock_Click()
If txtstocks5.Text = Empty Then
    txtstocks5.Text = "0"
    
    MsgBox "Stock Updated!", vbInformation, ""
    
    cmdSaveUpdateStock.Caption = "SAVED"
    cmdSaveUpdateStock.Enabled = False
    
    With frmMain.stocksdata.Recordset
        .Edit
        .Fields("stocks") = Val(txtstocks5.Text)
        .Update
    End With
    
    frmMain.stocksdata.RecordSource = "select * from tblstocks"
    frmMain.stocksdata.Refresh
    
    frmMain.Enabled = True
    frmMain.stocksdata.RecordSource = "select * from tblstocks"
    frmMain.stocksdata.Refresh
    frmMain.txtsearchproduct.Text = ""
    frmMain.txtsearchproduct.SetFocus
    frmMain.Show
    Unload Me
    
Else
    If Val(txtstocks5.Text) = 0 Then
        txtstocks5.Text = "0"
    
        MsgBox "Stock Updated!", vbInformation, ""
        
        cmdSaveUpdateStock.Caption = "SAVED"
        cmdSaveUpdateStock.Enabled = False
        
        With frmMain.stocksdata.Recordset
            .Edit
            .Fields("stocks") = Val(txtstocks5.Text)
            .Update
        End With
        
        frmMain.stocksdata.RecordSource = "select * from tblstocks"
        frmMain.stocksdata.Refresh
        
        frmMain.Enabled = True
        frmMain.stocksdata.RecordSource = "select * from tblstocks"
        frmMain.stocksdata.Refresh
        frmMain.txtsearchproduct.Text = ""
        frmMain.txtsearchproduct.SetFocus
        frmMain.Show
        Unload Me
    Else
        MsgBox "Stock Updated!", vbInformation, ""
        
        cmdSaveUpdateStock.Caption = "SAVED"
        cmdSaveUpdateStock.Enabled = False
        
        With frmMain.stocksdata.Recordset
            .Edit
            .Fields("stocks") = Val(txtstocks5.Text)
            .Update
        End With
        
        frmMain.Enabled = True
        frmMain.stocksdata.RecordSource = "select * from tblstocks"
        frmMain.stocksdata.Refresh
        frmMain.txtsearchproduct.Text = ""
        frmMain.txtsearchproduct.SetFocus
        frmMain.Show
        Unload Me
    End If
End If
End Sub

Private Sub Command1_Click()
Dim reply As String
reply = MsgBox("Close?", vbQuestion + vbYesNo, "")

If reply = vbYes Then
    frmMain.Enabled = True
    frmMain.productsdatainv.RecordSource = "select * from tblproducts"
    frmMain.productsdatainv.Refresh
    frmMain.txtsearchproduct_2.Text = ""
    frmMain.Show
    Unload Me
    frmMain.txtsearchproduct_2.SetFocus
End If
End Sub

Private Sub Form_Load()
productsdatainv.RecordSource = "select * from tblproducts'"
productsdatainv.Refresh

If Not productsdatainv.Recordset.EOF Then    '<<<<<<<---- NAGKAKAERROR KA NG DUPLICATION DAHIL KAPAG MAY NA DOBLENG NAME
    Do
    cmbProducts.AddItem productsdatainv.Recordset.Fields("barcode")
    productsdatainv.Recordset.MoveNext
    Loop Until productsdatainv.Recordset.EOF
End If
End Sub




Private Sub txtBarcode1_Change()
lblmessage3_3.Visible = False
End Sub

Private Sub txtBarcode1_GotFocus()
frmView.txtBarcode1.SelStart = Len(frmView.txtBarcode1.Text)
End Sub

Private Sub txtBarcode1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789.- $%&*", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtbarcode2_Change()
lblmessage4_3.Visible = False

If Not txtbarcode2.Text = frmMain.productsdatainv.Recordset.Fields("barcode") Or Not txtprodname2.Text = frmMain.productsdatainv.Recordset.Fields("productname") Or Not txtprice2.Text = frmMain.productsdatainv.Recordset.Fields("price") Or Not txtdescription2.Text = frmMain.productsdatainv.Recordset.Fields("description") Then
    cmdSaveUpdate.Caption = "SAVE CHANGES"
    cmdSaveUpdate.Enabled = True
Else
    cmdSaveUpdate.Enabled = False
End If
End Sub

Private Sub txtbarcode2_GotFocus()
frmView.txtbarcode2.SelStart = Len(frmView.txtbarcode2.Text)
End Sub

Private Sub txtbarcode2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789.- $%&*", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtdescription1_GotFocus()
frmView.txtdescription1.SelStart = Len(frmView.txtdescription1.Text)
End Sub

Private Sub txtdescription1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789!?.,()_-&%$#/ ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtdescription2_Change()
If Not txtbarcode2.Text = frmMain.productsdatainv.Recordset.Fields("barcode") Or Not txtprodname2.Text = frmMain.productsdatainv.Recordset.Fields("productname") Or Not txtprice2.Text = frmMain.productsdatainv.Recordset.Fields("price") Or Not txtdescription2.Text = frmMain.productsdatainv.Recordset.Fields("description") Then
    cmdSaveUpdate.Caption = "SAVE CHANGES"
    cmdSaveUpdate.Enabled = True
Else
    cmdSaveUpdate.Enabled = False
End If
End Sub

Private Sub txtdescription2_GotFocus()
frmView.txtdescription2.SelStart = Len(frmView.txtdescription2.Text)
End Sub

Private Sub txtdescription2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789!?.,()_-&%$#/ ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtprice1_Change()
lblmessage3_2.Visible = False
End Sub

Private Sub txtprice1_GotFocus()
frmView.txtprice1.SelStart = Len(frmView.txtprice1.Text)
End Sub

Private Sub txtprice1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtprice2_Change()
lblmessage4_2.Visible = False

If Not txtbarcode2.Text = frmMain.productsdatainv.Recordset.Fields("barcode") Or Not txtprodname2.Text = frmMain.productsdatainv.Recordset.Fields("productname") Or Not txtprice2.Text = frmMain.productsdatainv.Recordset.Fields("price") Or Not txtdescription2.Text = frmMain.productsdatainv.Recordset.Fields("description") Then
    cmdSaveUpdate.Caption = "SAVE CHANGES"
    cmdSaveUpdate.Enabled = True
Else
    cmdSaveUpdate.Enabled = False
End If
End Sub

Private Sub txtprice2_GotFocus()
    frmView.txtprice2.SelStart = Len(frmView.txtprice2.Text)
End Sub

Private Sub txtprice2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtprodname1_Change()
lblmessage3_1.Visible = False
End Sub

Private Sub txtprodname1_GotFocus()
frmView.txtprodname1.SelStart = Len(frmView.txtprodname1.Text)
End Sub

Private Sub txtprodname1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789()-_,.#!&/% ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtprodname2_Change()
lblmessage4_1.Visible = False

If Not txtbarcode2.Text = frmMain.productsdatainv.Recordset.Fields("barcode") Or Not txtprodname2.Text = frmMain.productsdatainv.Recordset.Fields("productname") Or Not txtprice2.Text = frmMain.productsdatainv.Recordset.Fields("price") Or Not txtdescription2.Text = frmMain.productsdatainv.Recordset.Fields("description") Then
    cmdSaveUpdate.Caption = "SAVE CHANGES"
    cmdSaveUpdate.Enabled = True
Else
    cmdSaveUpdate.Enabled = False
End If
End Sub

Private Sub txtprodname2_GotFocus()
    frmView.txtprodname2.SelStart = Len(frmView.txtprodname2.Text)
End Sub

Private Sub txtprodname2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789()-_,.#!&%/ ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtstocks4_Change()
lblmessage5_2.Visible = False
End Sub

Private Sub txtstocks4_GotFocus()
frmView.txtstocks4.SelStart = Len(frmView.txtstocks4.Text)
End Sub

Private Sub txtstocks4_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtstocks5_Change()
If txtstocks5.Text = frmMain.stocksdata.Recordset.Fields("stocks") Then
    cmdSaveUpdateStock.Caption = "SAVE CHANGES"
    cmdSaveUpdateStock.Enabled = False
Else
    cmdSaveUpdateStock.Caption = "SAVE CHANGES"
    cmdSaveUpdateStock.Enabled = True
End If
End Sub

Private Sub txtstocks5_GotFocus()
frmView.txtstocks5.SelStart = Len(frmView.txtstocks5.Text)
End Sub

Private Sub txtstocks5_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
