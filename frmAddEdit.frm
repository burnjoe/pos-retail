VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAddEdit 
   BackColor       =   &H003E472C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   4470
   ClientTop       =   4875
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMyAccount 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   120
      TabIndex        =   68
      Top             =   120
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Frame Frame2 
         BackColor       =   &H0082D1B0&
         Caption         =   "CHANGE PASSWORD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   85
         Top             =   600
         Width           =   11535
         Begin VB.CheckBox chkShow 
            BackColor       =   &H0082D1B0&
            Caption         =   "SHOW PASSWORD"
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
            Left            =   8280
            TabIndex        =   129
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   360
            TabIndex        =   123
            Top             =   840
            Width           =   4215
            Begin VB.Label Label34 
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
               TabIndex        =   125
               Top             =   240
               Width           =   1830
            End
            Begin VB.Label lblusername10 
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
               Left            =   2040
               TabIndex        =   124
               Top             =   240
               Width           =   105
            End
         End
         Begin VB.TextBox txtconfirmpass 
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
            IMEMode         =   3  'DISABLE
            Left            =   4800
            MaxLength       =   35
            PasswordChar    =   "•"
            TabIndex        =   3
            Top             =   2175
            Width           =   3135
         End
         Begin VB.TextBox txtchangepass 
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
            IMEMode         =   3  'DISABLE
            Left            =   4800
            MaxLength       =   35
            PasswordChar    =   "•"
            TabIndex        =   2
            Top             =   1440
            Width           =   3135
         End
         Begin VB.TextBox txtcurrentpass 
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
            IMEMode         =   3  'DISABLE
            Left            =   4800
            MaxLength       =   35
            PasswordChar    =   "•"
            TabIndex        =   1
            Top             =   600
            Width           =   3135
         End
         Begin VB.CommandButton cmdChangePass 
            BackColor       =   &H00DBF2E9&
            Caption         =   "CHANGE PASSWORD"
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
            Left            =   9120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label lblmessage_11 
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
            Left            =   6720
            TabIndex        =   128
            Top             =   1920
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage_10 
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
            Left            =   6600
            TabIndex        =   127
            Top             =   1200
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage_9 
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
            Left            =   6720
            TabIndex        =   126
            Top             =   360
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CONFIRM PASSWORD:"
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
            Left            =   4800
            TabIndex        =   107
            Top             =   1920
            Width           =   1800
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CHANGE PASSWORD:"
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
            Left            =   4800
            TabIndex        =   106
            Top             =   1200
            Width           =   1680
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CURRENT PASSWORD:"
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
            Left            =   4800
            TabIndex        =   105
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label lblusertype10 
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
            Left            =   1440
            TabIndex        =   96
            Top             =   1680
            Width           =   105
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "USER TYPE:"
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
            TabIndex        =   95
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label lblid10 
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
            Left            =   1680
            TabIndex        =   87
            Top             =   600
            Width           =   105
         End
         Begin VB.Label Label33 
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
            TabIndex        =   86
            Top             =   600
            Width           =   1080
         End
      End
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
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0082D1B0&
         Caption         =   "PERSONAL INFORMATION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   69
         Top             =   3600
         Width           =   11535
         Begin VB.Frame Frame8 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            TabIndex        =   118
            Top             =   1920
            Width           =   10935
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ADDRESS:"
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
               TabIndex        =   122
               Top             =   120
               Width           =   795
            End
            Begin VB.Label lbladdress10 
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
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   7680
            TabIndex        =   117
            Top             =   1080
            Width           =   3495
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "EMAIL:"
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
               TabIndex        =   120
               Top             =   120
               Width           =   570
            End
            Begin VB.Label lblemail10 
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
               TabIndex        =   119
               Top             =   360
               Width           =   45
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   7680
            TabIndex        =   114
            Top             =   240
            Width           =   3495
            Begin VB.Label Label41 
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
               TabIndex        =   116
               Top             =   120
               Width           =   1005
            End
            Begin VB.Label lbllast10 
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
               TabIndex        =   115
               Top             =   360
               Width           =   45
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            TabIndex        =   111
            Top             =   240
            Width           =   3615
            Begin VB.Label Label39 
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
               TabIndex        =   113
               Top             =   120
               Width           =   1035
            End
            Begin VB.Label lblfirst10 
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
               TabIndex        =   112
               Top             =   360
               Width           =   45
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H0082D1B0&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   3960
            TabIndex        =   108
            Top             =   240
            Width           =   3615
            Begin VB.Label Label40 
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
               TabIndex        =   110
               Top             =   120
               Width           =   1245
            End
            Begin VB.Label lblmiddle10 
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
               TabIndex        =   109
               Top             =   360
               Width           =   45
            End
         End
         Begin VB.Label lblphone10 
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
            Left            =   5760
            TabIndex        =   104
            Top             =   1440
            Width           =   45
         End
         Begin VB.Label lblbday10 
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
            Left            =   3480
            TabIndex        =   103
            Top             =   1440
            Width           =   45
         End
         Begin VB.Label lblsex10 
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
            Left            =   1920
            TabIndex        =   102
            Top             =   1440
            Width           =   45
         End
         Begin VB.Label Label49 
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
            Left            =   5640
            TabIndex        =   101
            Top             =   1200
            Width           =   645
         End
         Begin VB.Label Label48 
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
            Left            =   3360
            TabIndex        =   100
            Top             =   1200
            Width           =   870
         End
         Begin VB.Label Label47 
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
            Left            =   1800
            TabIndex        =   99
            Top             =   1200
            Width           =   345
         End
         Begin VB.Label Label46 
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
            Left            =   360
            TabIndex        =   98
            Top             =   1200
            Width           =   390
         End
         Begin VB.Label lblage10 
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
            Left            =   480
            TabIndex        =   97
            Top             =   1440
            Width           =   45
         End
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   4635
         Picture         =   "frmAddEdit.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MY ACCOUNT"
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
         Left            =   5250
         TabIndex        =   70
         Top             =   120
         Width           =   1650
      End
   End
   Begin VB.Frame frmUpdateAccount 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      Height          =   6615
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Frame Frame1 
         BackColor       =   &H0082D1B0&
         Caption         =   "PERSONAL INFORMATION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   120
         TabIndex        =   67
         Top             =   720
         Width           =   11535
         Begin VB.CommandButton cmdstatus 
            BackColor       =   &H00C0C0FF&
            Caption         =   "DEACTIVATE "
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
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   4800
            Width           =   1815
         End
         Begin VB.CommandButton cmdsave2 
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
            Left            =   9480
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   4800
            Width           =   1695
         End
         Begin VB.ComboBox cmbUserType2 
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
            ItemData        =   "frmAddEdit.frx":1E6F3
            Left            =   240
            List            =   "frmAddEdit.frx":1E6FD
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1560
            Width           =   2535
         End
         Begin VB.ComboBox cmbSex2 
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
            ItemData        =   "frmAddEdit.frx":1E70E
            Left            =   3000
            List            =   "frmAddEdit.frx":1E718
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox txtFirst2 
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
            MaxLength       =   50
            TabIndex        =   21
            Top             =   2400
            Width           =   3015
         End
         Begin VB.TextBox txtMiddle2 
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
            Left            =   3480
            MaxLength       =   50
            TabIndex        =   22
            Top             =   2400
            Width           =   3015
         End
         Begin VB.TextBox txtLast2 
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
            Left            =   6720
            MaxLength       =   50
            TabIndex        =   23
            Top             =   2400
            Width           =   3015
         End
         Begin VB.TextBox txtAge2 
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
            Left            =   9960
            MaxLength       =   3
            TabIndex        =   24
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtAddress2 
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
            Left            =   4680
            MaxLength       =   200
            TabIndex        =   27
            Top             =   3120
            Width           =   6615
         End
         Begin VB.TextBox txtPhone2 
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
            MaxLength       =   15
            TabIndex        =   28
            Top             =   3840
            Width           =   3255
         End
         Begin VB.TextBox txtEmail2 
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
            MaxLength       =   50
            TabIndex        =   29
            Top             =   3840
            Width           =   3375
         End
         Begin MSComCtl2.DTPicker dtbday2 
            Height          =   330
            Left            =   240
            TabIndex        =   25
            Top             =   3120
            Width           =   2535
            _ExtentX        =   4471
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
            Format          =   107020291
            CurrentDate     =   43880
         End
         Begin VB.Label lblmessage2_7 
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
            TabIndex        =   94
            Top             =   1320
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage2_6 
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
            TabIndex        =   93
            Top             =   3600
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage2_4 
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
            Left            =   5640
            TabIndex        =   92
            Top             =   2880
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage2_5 
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
            Left            =   1800
            TabIndex        =   91
            Top             =   3600
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage2_3 
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
            Left            =   10560
            TabIndex        =   90
            Top             =   2160
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage2_2 
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
            Left            =   7920
            TabIndex        =   89
            Top             =   2160
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage2_1 
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
            Left            =   1440
            TabIndex        =   88
            Top             =   2160
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblusername_2 
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
            Left            =   2280
            TabIndex        =   84
            Top             =   840
            Width           =   105
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT USERNAME:"
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
            TabIndex        =   83
            Top             =   840
            Width           =   1755
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* USER TYPE:"
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
            TabIndex        =   82
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Label lblaccid_2 
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
            Left            =   1560
            TabIndex        =   81
            Top             =   480
            Width           =   105
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT ID:"
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
            TabIndex        =   80
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* SEX:"
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
            Left            =   3000
            TabIndex        =   79
            Top             =   2880
            Width           =   465
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* FIRST NAME:"
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
            TabIndex        =   78
            Top             =   2160
            Width           =   1140
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MIDDLE NAME:"
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
            Left            =   3480
            TabIndex        =   77
            Top             =   2160
            Width           =   1200
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* LAST NAME:"
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
            Left            =   6720
            TabIndex        =   76
            Top             =   2160
            Width           =   1065
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* AGE:"
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
            Left            =   9960
            TabIndex        =   75
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* ADDRESS:"
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
            Left            =   4680
            TabIndex        =   74
            Top             =   2880
            Width           =   900
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* BIRTHDAY:"
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
            TabIndex        =   73
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* PHONE NUMBER:"
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
            TabIndex        =   72
            Top             =   3600
            Width           =   1500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* EMAIL ADDRESS:"
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
            Left            =   3720
            TabIndex        =   71
            Top             =   3600
            Width           =   1440
         End
      End
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
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   4320
         Picture         =   "frmAddEdit.frx":1E72A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UPDATE ACCOUNT"
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
         Left            =   4905
         TabIndex        =   66
         Top             =   120
         Width           =   2205
      End
   End
   Begin VB.Frame frmAddAccount 
      BackColor       =   &H0082D1B0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Frame frmPersonalInfo 
         BackColor       =   &H0082D1B0&
         Caption         =   "PERSONAL INFORMATION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   43
         Top             =   3360
         Width           =   11535
         Begin VB.CommandButton cmdCreateAccount 
            BackColor       =   &H00DBF2E9&
            Caption         =   "CREATE ACCOUNT"
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
            Left            =   9480
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txtEmail1 
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
            MaxLength       =   50
            TabIndex        =   18
            Top             =   2160
            Width           =   3375
         End
         Begin VB.TextBox txtPhone1 
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
            MaxLength       =   15
            TabIndex        =   17
            Top             =   2160
            Width           =   3255
         End
         Begin VB.TextBox txtAddress1 
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
            Left            =   4680
            MaxLength       =   200
            TabIndex        =   16
            Top             =   1440
            Width           =   6615
         End
         Begin VB.TextBox txtAge1 
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
            Left            =   9960
            MaxLength       =   3
            TabIndex        =   13
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtLast1 
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
            Left            =   6720
            MaxLength       =   50
            TabIndex        =   12
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtMiddle1 
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
            Left            =   3480
            MaxLength       =   50
            TabIndex        =   11
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtFirst1 
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
            MaxLength       =   50
            TabIndex        =   10
            Top             =   720
            Width           =   3015
         End
         Begin VB.ComboBox cmbsex1 
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
            ItemData        =   "frmAddEdit.frx":43A55
            Left            =   3000
            List            =   "frmAddEdit.frx":43A5F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1440
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker dtBday1 
            Height          =   330
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   2535
            _ExtentX        =   4471
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
            Format          =   107020291
            CurrentDate     =   43880
         End
         Begin VB.Label lblmessage1_12 
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
            TabIndex        =   65
            Top             =   1920
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage1_11 
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
            Left            =   3600
            TabIndex        =   63
            Top             =   1200
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage1_10 
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
            Left            =   10560
            TabIndex        =   62
            Top             =   480
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage1_8 
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
            Left            =   1800
            TabIndex        =   60
            Top             =   1920
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage1_6 
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
            Left            =   7920
            TabIndex        =   59
            Top             =   480
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage1_5 
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
            Left            =   4800
            TabIndex        =   58
            Top             =   480
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage1_4 
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
            Left            =   1440
            TabIndex        =   57
            Top             =   480
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* EMAIL ADDRESS:"
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
            Left            =   3720
            TabIndex        =   53
            Top             =   1920
            Width           =   1440
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* PHONE NUMBER:"
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
            TabIndex        =   52
            Top             =   1920
            Width           =   1500
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* BIRTHDAY:"
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
            TabIndex        =   51
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* ADDRESS:"
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
            Left            =   4680
            TabIndex        =   50
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* AGE:"
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
            Left            =   9960
            TabIndex        =   49
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* LAST NAME:"
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
            Left            =   6720
            TabIndex        =   48
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MIDDLE NAME:"
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
            Left            =   3480
            TabIndex        =   47
            Top             =   480
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* FIRST NAME:"
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
            TabIndex        =   46
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* SEX:"
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
            Left            =   3000
            TabIndex        =   45
            Top             =   1200
            Width           =   465
         End
         Begin VB.Label lblmessage1_7 
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
            Left            =   5760
            TabIndex        =   44
            Top             =   1200
            Visible         =   0   'False
            Width           =   45
         End
      End
      Begin VB.Frame frmAccountInfo 
         BackColor       =   &H0082D1B0&
         Caption         =   "ACCOUNT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   11535
         Begin VB.CheckBox chkshowpass1 
            BackColor       =   &H0082D1B0&
            Caption         =   "SHOW PASSWORD"
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
            Left            =   7680
            TabIndex        =   64
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Data accountsdata 
            Caption         =   "ACCOUNTS"
            Connect         =   "Access"
            DatabaseName    =   "database\database_sabana.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   8880
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "tblaccounts"
            Top             =   240
            Visible         =   0   'False
            Width           =   2580
         End
         Begin VB.ComboBox cmbUserType1 
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
            ItemData        =   "frmAddEdit.frx":43A71
            Left            =   360
            List            =   "frmAddEdit.frx":43A7B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtUsername1 
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
            Left            =   4260
            MaxLength       =   35
            TabIndex        =   7
            Top             =   600
            Width           =   3135
         End
         Begin VB.TextBox txtPassword1 
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
            IMEMode         =   3  'DISABLE
            Left            =   4260
            MaxLength       =   35
            PasswordChar    =   "•"
            TabIndex        =   8
            Top             =   1320
            Width           =   3135
         End
         Begin VB.TextBox txtretypepassword1 
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
            IMEMode         =   3  'DISABLE
            Left            =   4260
            MaxLength       =   35
            PasswordChar    =   "•"
            TabIndex        =   9
            Top             =   2040
            Width           =   3135
         End
         Begin VB.Label lblmessage1_9 
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
            Left            =   1440
            TabIndex        =   61
            Top             =   840
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACCOUNT ID:"
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
            TabIndex        =   56
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label lblaccid1 
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
            Left            =   1680
            TabIndex        =   55
            Top             =   360
            Width           =   105
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* USER TYPE:"
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
            TabIndex        =   54
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* CREATE USERNAME:"
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
            Left            =   4260
            TabIndex        =   42
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* CREATE PASSWORD:"
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
            Left            =   4260
            TabIndex        =   41
            Top             =   1080
            Width           =   1740
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* RETYPE PASSWORD:"
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
            Left            =   4260
            TabIndex        =   40
            Top             =   1800
            Width           =   1725
         End
         Begin VB.Label lblmessage1_1 
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
            Left            =   6120
            TabIndex        =   39
            Top             =   360
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage1_2 
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
            Left            =   6120
            TabIndex        =   38
            Top             =   1080
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblmessage1_3 
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
            Left            =   6120
            TabIndex        =   37
            Top             =   1800
            Visible         =   0   'False
            Width           =   45
         End
      End
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
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   3960
         Picture         =   "frmAddEdit.frx":43A8C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CREATE NEW ACCOUNT"
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
         Left            =   4612
         TabIndex        =   34
         Top             =   120
         Width           =   2790
      End
   End
End
Attribute VB_Name = "frmAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dtbday As String
Public accid As Integer

Private Sub Check1_Change()
txtPassword1.PasswordChar = ""
txtretypepassword1.PasswordChar = ""
End Sub

Private Sub chkShow_Click()
If chkShow.Value = 1 Then
    txtchangepass.PasswordChar = ""
    txtconfirmpass.PasswordChar = ""
Else
    txtchangepass.PasswordChar = "•"
    txtconfirmpass.PasswordChar = "•"
End If
End Sub

Private Sub chkshowpass1_Click()
If chkshowpass1.Value = 1 Then
    txtPassword1.PasswordChar = ""
    txtretypepassword1.PasswordChar = ""
Else
    txtPassword1.PasswordChar = "•"
    txtretypepassword1.PasswordChar = "•"
End If
End Sub

Private Sub cmbsex1_click()
lblmessage1_11.Visible = False
End Sub

Private Sub cmbSex2_Click()
dtbday = Format(dtbday2, "mmmm dd, yyyy")

If Not cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype") Or Not txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first") Or Not txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle") Or Not txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last") Or Not txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age") Or Not cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex") Or Not txtAddress2.Text = frmMain.accountsdata.Recordset.Fields("address") Or Not txtPhone2.Text = frmMain.accountsdata.Recordset.Fields("phone") Or Not txtEmail2.Text = frmMain.accountsdata.Recordset.Fields("email") Then
    cmdsave2.Enabled = True
    cmdsave2.Caption = "SAVE CHANGES"
Else
    cmdsave2.Enabled = False
    If dtbday = frmMain.accountsdata.Recordset.Fields("bday") Then
        cmdsave2.Enabled = False
    Else
        cmdsave2.Caption = "SAVE CHANGES"
        cmdsave2.Enabled = True
    End If
End If

End Sub

Private Sub cmbUserType1_Click()
    lblmessage1_9.Visible = False
End Sub

Private Sub cmbUserType2_Click()
dtbday = Format(dtbday2, "mmmm dd, yyyy")

If Not cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype") Or Not txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first") Or Not txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle") Or Not txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last") Or Not txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age") Or Not cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex") Or Not txtAddress2.Text = frmMain.accountsdata.Recordset.Fields("address") Or Not txtPhone2.Text = frmMain.accountsdata.Recordset.Fields("phone") Or Not txtEmail2.Text = frmMain.accountsdata.Recordset.Fields("email") Then
    cmdsave2.Enabled = True
    cmdsave2.Caption = "SAVE CHANGES"
Else
    cmdsave2.Enabled = False
    If dtbday = frmMain.accountsdata.Recordset.Fields("bday") Then
        cmdsave2.Enabled = False
    Else
        cmdsave2.Caption = "SAVE CHANGES"
        cmdsave2.Enabled = True
    End If
End If

End Sub

Private Sub cmdCancel1_Click()
Dim reply As String

    reply = MsgBox("Close?", vbQuestion + vbYesNo, "")
        
If reply = vbYes Then
    Unload Me
    frmMain.Show
    frmMain.Enabled = True
    frmMain.cmbUserType = "GENERAL"
    frmMain.cmbAccountStatus = "GENERAL"
    frmMain.txtsearchaccount.Text = ""
    frmMain.cmdAddAdmin.BackColor = &H82D1B0
End If


End Sub

Private Sub cmdCancel2_Click()
Dim reply As String

reply = MsgBox("Close?", vbQuestion + vbYesNo, "")

If reply = vbYes Then
    frmMain.Enabled = True
    frmMain.cmbUserType = "GENERAL"
    frmMain.cmbAccountStatus = "GENERAL"
    frmMain.cmdUpdateAdmin.BackColor = &H82D1B0
    frmMain.txtsearchaccount.Text = ""
    frmMain.Show
    Unload Me
End If

End Sub

Private Sub cmdCancel3_Click()
Dim reply As String

reply = MsgBox("Close?", vbQuestion + vbYesNo, "")

If reply = vbYes Then
    Unload Me
    frmMain.Show
    frmMain.Enabled = True
    frmMain.cmdMyAccount.BackColor = &H82D1B0
End If
End Sub

Private Sub cmdChangePass_Click()
If Not frmLogin.txtUsername1.Text = Empty Then 'para sa admin accounts
    frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername1.Text + "' and pass= '" + txtcurrentpass.Text + "'"
    frmLogin.accountsdata.Refresh
    
    If txtcurrentpass.Text = Empty Then
        lblmessage_9.Caption = "(Enter Current Password)"
        lblmessage_9.Visible = True
        txtcurrentpass.SetFocus
    Else
        If frmLogin.accountsdata.Recordset.EOF Then
            lblmessage_9.Caption = "(Incorrect Password)"
            lblmessage_9.Visible = True
            txtcurrentpass.SetFocus
        Else
            If Len(txtchangepass) < 4 Then
                lblmessage_10.Caption = "(Enter At Least 4 Characters Long)"
                lblmessage_10.Visible = True
                txtchangepass.SetFocus
            Else
                If Len(txtconfirmpass) < 4 Then
                    lblmessage_11.Caption = "(Please Confirm Password)"
                    lblmessage_11.Visible = True
                    txtconfirmpass.SetFocus
                Else
                    If Not txtchangepass.Text = txtconfirmpass.Text Then
                        lblmessage_11.Caption = "(Password Did Not Match)"
                        lblmessage_11.Visible = True
                        txtconfirmpass.SetFocus
                    Else
                        If txtchangepass.Text = frmLogin.accountsdata.Recordset.Fields("pass") Then
                            lblmessage_10.Caption = "(Enter New Password)"
                            lblmessage_10.Visible = True
                            txtchangepass.SetFocus
                            txtconfirmpass.Text = ""
                        Else
                            MsgBox "Password Changed!", vbInformation, ""
                            With frmLogin.accountsdata.Recordset
                                .Edit
                                .Fields("pass") = txtchangepass.Text
                                .Update
                            End With
                            
                            frmLogin.txtPassword1.Text = txtchangepass.Text
                            
                            txtcurrentpass.Text = ""
                            txtchangepass.Text = ""
                            txtconfirmpass.Text = ""
                            txtcurrentpass.SetFocus
                            chkShow.Value = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
Else
    frmLogin.accountsdata.RecordSource = "select * from tblaccounts where user= '" + frmLogin.txtUsername2.Text + "' and pass= '" + txtcurrentpass.Text + "'"
    frmLogin.accountsdata.Refresh
    
    If txtcurrentpass.Text = Empty Then
        lblmessage_9.Caption = "(Enter Current Password)"
        lblmessage_9.Visible = True
        txtcurrentpass.SetFocus
    Else
        If frmLogin.accountsdata.Recordset.EOF Then
            lblmessage_9.Caption = "(Incorrect Password)"
            lblmessage_9.Visible = True
            txtcurrentpass.SetFocus
        Else
            If Len(txtchangepass) < 4 Then
                lblmessage_10.Caption = "(Enter At Least 4 Characters Long)"
                lblmessage_10.Visible = True
                txtchangepass.SetFocus
            Else
                If Len(txtconfirmpass) < 4 Then
                    lblmessage_11.Caption = "(Please Confirm Password)"
                    lblmessage_11.Visible = True
                    txtconfirmpass.SetFocus
                Else
                    If Not txtchangepass.Text = txtconfirmpass.Text Then
                        lblmessage_11.Caption = "(Password Did Not Match)"
                        lblmessage_11.Visible = True
                        txtconfirmpass.SetFocus
                    Else
                        If txtchangepass.Text = frmLogin.accountsdata.Recordset.Fields("pass") Then
                            lblmessage_10.Caption = "(Enter New Password)"
                            lblmessage_10.Visible = True
                            txtchangepass.SetFocus
                            txtconfirmpass.Text = ""
                        Else
                            MsgBox "Password Changed!", vbInformation, ""
                            With frmLogin.accountsdata.Recordset
                                .Edit
                                .Fields("pass") = txtchangepass.Text
                                .Update
                            End With
                            
                            frmLogin.txtPassword2.Text = txtchangepass.Text
                            
                            txtcurrentpass.Text = ""
                            txtchangepass.Text = ""
                            txtconfirmpass.Text = ""
                            txtcurrentpass.SetFocus
                            chkShow.Value = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub cmdCreateAccount_Click()

'to analyze what to check before adding new account in database
If cmbUserType1 = Empty Then
    lblmessage1_9.Caption = "(Specify)"
    lblmessage1_9.Visible = True
    cmbUserType1.SetFocus
Else
    If Len(txtUsername1) < 4 Then
        lblmessage1_1.Caption = "(Enter At Least 4 Characters Long)"
        lblmessage1_1.Visible = True
        txtUsername1.SetFocus
    Else
        If Len(txtPassword1) < 4 Then
            lblmessage1_2.Caption = "(Enter At Least 4 Characters Long)"
            lblmessage1_2.Visible = True
            txtPassword1.SetFocus
        Else
            If txtretypepassword1.Text = Empty Then
                lblmessage1_3.Caption = "(Please Retype Password)"
                lblmessage1_3.Visible = True
                txtretypepassword1.SetFocus
            Else
                If Len(txtFirst1) < 2 Then
                    lblmessage1_4.Caption = "(Enter Valid Name)"
                    lblmessage1_4.Visible = True
                    txtFirst1.SetFocus
                Else
                    If Len(txtLast1) < 2 Then
                        lblmessage1_6.Caption = "(Enter Valid Name)"
                        lblmessage1_6.Visible = True
                        txtLast1.SetFocus
                    Else
                        If txtAge1.Text = Empty Then
                            lblmessage1_10.Caption = "(Enter Age)"
                            lblmessage1_10.Visible = True
                            txtAge1.SetFocus
                        Else
                            If cmbsex1 = Empty Then
                                lblmessage1_11.Caption = "(Specify)"
                                lblmessage1_11.Visible = True
                                cmbsex1.SetFocus
                            Else
                                If Len(txtAddress1) < 2 Then
                                    lblmessage1_7.Caption = "(Enter Valid Address)"
                                    lblmessage1_7.Visible = True
                                    txtAddress1.SetFocus
                                Else
                                    If Len(txtPhone1) < 11 Then
                                        lblmessage1_8.Caption = "(Enter Valid Contact)"
                                        lblmessage1_8.Visible = True
                                        txtPhone1.SetFocus
                                    Else
                                        If Len(txtEmail1) < 10 Then
                                            lblmessage1_12.Caption = "(Enter Valid Email Address)"
                                            lblmessage1_12.Visible = True
                                            txtEmail1.SetFocus
                                        Else
                                            accountsdata.RecordSource = "select * from tblaccounts where user= '" + txtUsername1.Text + "'"
                                            accountsdata.Refresh
                                            
                                            If accountsdata.Recordset.EOF Then
                                                If txtPassword1.Text = txtretypepassword1.Text Then
                                                    With accountsdata.Recordset
                                                        .AddNew
                                                        .Fields("accid") = lblaccid1.Caption
                                                        .Fields("usertype") = cmbUserType1
                                                        .Fields("user") = txtUsername1.Text
                                                        .Fields("pass") = txtPassword1.Text
                                                        .Fields("first") = txtFirst1.Text
                                                        .Fields("middle") = txtMiddle1.Text
                                                        .Fields("last") = txtLast1.Text
                                                        .Fields("age") = txtAge1.Text
                                                        .Fields("sex") = cmbsex1
                                                        .Fields("bday") = Format(dtBday1, "mmmm dd, yyyy")
                                                        .Fields("address") = txtAddress1.Text
                                                        .Fields("phone") = txtPhone1.Text
                                                        .Fields("email") = txtEmail1.Text
                                                        .Fields("datecreated") = Format(Date, "mmmm dd, yyyy")
                                                        .Fields("timecreated") = Time
                                                        .Fields("status") = "ACTIVE"
                                                        .Update
                                                    End With
                                                    
                                                    MsgBox "Account Registered Successfully!", vbInformation, ""
                                                    
                                                    accountsdata.RecordSource = "select * from tblaccounts"
                                                    accountsdata.Refresh
                                                    
                                                    frmMain.accountsdata.Refresh
                                                     
                                                    Do
                                                    accid = frmAddEdit.accountsdata.Recordset.Fields("accid")
                                                    accountsdata.Recordset.MoveNext
                                                    Loop Until accountsdata.Recordset.EOF
                                                    
                                                    accid = frmAddEdit.accid + 1
                                                    lblaccid1.Caption = accid
                                                    
                                                    'clear all
                                                    txtUsername1.Text = ""
                                                    txtPassword1.Text = ""
                                                    txtretypepassword1.Text = ""
                                                    txtFirst1.Text = ""
                                                    txtMiddle1.Text = ""
                                                    txtLast1.Text = ""
                                                    txtAge1.Text = ""
                                                    txtAddress1.Text = ""
                                                    txtPhone1.Text = ""
                                                    txtEmail1.Text = ""
                                                    chkshowpass1 = 0
                                                    
                                                    'dito yung end
                                                Else
                                                    lblmessage1_3.Caption = "(Password Did Not Match)"
                                                    lblmessage1_3.Visible = True
                                                    txtretypepassword1.SetFocus
                                                End If
                                            Else
                                                lblmessage1_1.Caption = "(Username Already Taken)"
                                                lblmessage1_1.Visible = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
End Sub



Private Sub cmdsave2_Click()

If Len(txtFirst2) < 2 Then
    lblmessage2_1.Caption = "(Enter Valid Name)"
    lblmessage2_1.Visible = True
    txtFirst2.SetFocus
Else
    If Len(txtLast2) < 2 Then
        lblmessage2_2.Caption = "(Enter Valid Name)"
        lblmessage2_2.Visible = True
        txtLast2.SetFocus
    Else
        If txtAge2.Text = Empty Then
            lblmessage2_3.Caption = "(Enter Age)"
            lblmessage2_3.Visible = True
            txtAge2.SetFocus
        Else
            If Len(txtAddress2) < 2 Then
                lblmessage2_4.Caption = "(Enter Valid Address)"
                lblmessage2_4.Visible = True
                txtAddress2.SetFocus
            Else
                If Len(txtPhone2) < 11 Then
                    lblmessage2_5.Caption = "(Enter Valid Contact)"
                    lblmessage2_5.Visible = True
                    txtPhone2.SetFocus
                Else
                    If Len(txtEmail2) < 10 Then
                        lblmessage2_6.Caption = "(Enter Valid Email Address)"
                        lblmessage2_6.Visible = True
                        txtEmail2.SetFocus
                    Else
                        
                        With frmMain.accountsdata.Recordset
                            .Edit
                            .Fields("usertype") = cmbUserType2
                            .Fields("first") = txtFirst2.Text
                            .Fields("middle") = txtMiddle2.Text
                            .Fields("last") = txtLast2.Text
                            .Fields("age") = txtAge2.Text
                            .Fields("bday") = Format(dtbday2, "mmmm dd, yyyy")
                            .Fields("sex") = cmbSex2
                            .Fields("address") = txtAddress2.Text
                            .Fields("phone") = txtPhone2.Text
                            .Fields("email") = txtEmail2.Text
                            .Update
                        End With
                        
                        If frmMain.accountsdata.Recordset.Fields("user") = frmLogin.accountsdata.Recordset.Fields("user") Then
                            frmMain.lblfirst.Caption = frmMain.accountsdata.Recordset.Fields("first")
                            frmMain.lbllast.Caption = frmMain.accountsdata.Recordset.Fields("last")
                        End If
                        
                        MsgBox "Account Updated!", vbInformation, ""
                        
                        cmdsave2.Enabled = False
                        cmdsave2.Caption = "SAVED"
                    End If
                End If
            End If
        End If
    End If
End If
End Sub

Private Sub cmdstatus_Click()
Dim reply As String

If cmdstatus.Caption = "DEACTIVATE" Then
    reply = MsgBox("Deactivate This Account?", vbQuestion + vbYesNo, "")
    
    If reply = vbYes Then
        frmMain.accountsdata.Recordset.Edit
        frmMain.accountsdata.Recordset.Fields("status") = "NOT ACTIVE"
        frmMain.accountsdata.Recordset.Update
        cmdstatus.Caption = "ACTIVATE"
        frmAddEdit.cmdstatus.BackColor = &HC0FFC0
    End If
Else
    reply = MsgBox("Activate This Account?", vbQuestion + vbYesNo, "")
    
    If reply = vbYes Then
        frmMain.accountsdata.Recordset.Edit
        frmMain.accountsdata.Recordset.Fields("status") = "ACTIVE"
        frmMain.accountsdata.Recordset.Update
        cmdstatus.Caption = "DEACTIVATE"
        frmAddEdit.cmdstatus.BackColor = &HC0C0FF
    End If
End If

End Sub

Private Sub dtbday2_Change()
dtbday = Format(dtbday2, "mmmm dd, yyyy")

If Not cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype") Or Not txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first") Or Not txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle") Or Not txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last") Or Not txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age") Or Not cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex") Or Not txtAddress2.Text = frmMain.accountsdata.Recordset.Fields("address") Or Not txtPhone2.Text = frmMain.accountsdata.Recordset.Fields("phone") Or Not txtEmail2.Text = frmMain.accountsdata.Recordset.Fields("email") Then
    cmdsave2.Enabled = True
    cmdsave2.Caption = "SAVE CHANGES"
Else
    cmdsave2.Enabled = False
    If dtbday = frmMain.accountsdata.Recordset.Fields("bday") Then
        cmdsave2.Enabled = False
    Else
        cmdsave2.Caption = "SAVE CHANGES"
        cmdsave2.Enabled = True
    End If
End If

End Sub

Private Sub Form_Load() ' error

'If frmAskAdminPassword.whatbutton = "addadmin" Then
   ' admindata.RecordSource = "select * from tbladmin"
   ' admindata.Refresh
   ' accid = admindata.Recordset.FindLast()
'End If

End Sub

Private Sub txtAddress1_Change()
lblmessage1_7.Visible = False
End Sub

Private Sub txtAddress1_GotFocus()
frmAddEdit.txtAddress1.SelStart = Len(frmAddEdit.txtAddress1.Text)
End Sub

Private Sub txtAddress1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ01234567890-.,_#() ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtAddress2_Change()
lblmessage2_4.Visible = False
dtbday = Format(dtbday2, "mmmm dd, yyyy")

If Not cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype") Or Not txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first") Or Not txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle") Or Not txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last") Or Not txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age") Or Not cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex") Or Not txtAddress2.Text = frmMain.accountsdata.Recordset.Fields("address") Or Not txtPhone2.Text = frmMain.accountsdata.Recordset.Fields("phone") Or Not txtEmail2.Text = frmMain.accountsdata.Recordset.Fields("email") Then
    cmdsave2.Enabled = True
    cmdsave2.Caption = "SAVE CHANGES"
Else
    cmdsave2.Enabled = False
    If dtbday = frmMain.accountsdata.Recordset.Fields("bday") Then
        cmdsave2.Enabled = False
    Else
        cmdsave2.Caption = "SAVE CHANGES"
        cmdsave2.Enabled = True
    End If
End If

End Sub

Private Sub txtAddress2_GotFocus()
frmAddEdit.txtAddress2.SelStart = Len(frmAddEdit.txtAddress2.Text)
End Sub

Private Sub txtAddress2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ01234567890-.,_#() ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtAge1_Change()
lblmessage1_10.Visible = False
End Sub

Private Sub txtAge1_GotFocus()
frmAddEdit.txtAge1.SelStart = Len(frmAddEdit.txtAge1.Text)
End Sub

Private Sub txtAge1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtAge2_Change()
lblmessage2_3.Visible = False
dtbday = Format(dtbday2, "mmmm dd, yyyy")

If Not cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype") Or Not txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first") Or Not txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle") Or Not txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last") Or Not txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age") Or Not cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex") Or Not txtAddress2.Text = frmMain.accountsdata.Recordset.Fields("address") Or Not txtPhone2.Text = frmMain.accountsdata.Recordset.Fields("phone") Or Not txtEmail2.Text = frmMain.accountsdata.Recordset.Fields("email") Then
    cmdsave2.Enabled = True
    cmdsave2.Caption = "SAVE CHANGES"
Else
    cmdsave2.Enabled = False
    If dtbday = frmMain.accountsdata.Recordset.Fields("bday") Then
        cmdsave2.Enabled = False
    Else
        cmdsave2.Caption = "SAVE CHANGES"
        cmdsave2.Enabled = True
    End If
End If

End Sub

Private Sub txtAge2_GotFocus()
frmAddEdit.txtAge2.SelStart = Len(frmAddEdit.txtAge2.Text)
End Sub

Private Sub txtAge2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtchangepass_Change()
lblmessage_10.Visible = False
End Sub

Private Sub txtchangepass_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789._", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtconfirmpass_Change()
lblmessage_11.Visible = False
End Sub

Private Sub txtconfirmpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789._", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtcurrentpass_Change()
lblmessage_9.Visible = False
End Sub

Private Sub txtcurrentpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789._", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtEmail1_Change()
lblmessage1_12.Visible = False
End Sub

Private Sub txtEmail1_GotFocus()
frmAddEdit.txtEmail1.SelStart = Len(frmAddEdit.txtEmail1.Text)
End Sub

Private Sub txtEmail1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789-_@., ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtEmail2_Change()
lblmessage2_6.Visible = False
dtbday = Format(dtbday2, "mmmm dd, yyyy")

If Not cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype") Or Not txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first") Or Not txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle") Or Not txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last") Or Not txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age") Or Not cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex") Or Not txtAddress2.Text = frmMain.accountsdata.Recordset.Fields("address") Or Not txtPhone2.Text = frmMain.accountsdata.Recordset.Fields("phone") Or Not txtEmail2.Text = frmMain.accountsdata.Recordset.Fields("email") Then
    cmdsave2.Enabled = True
    cmdsave2.Caption = "SAVE CHANGES"
Else
    cmdsave2.Enabled = False
    If dtbday = frmMain.accountsdata.Recordset.Fields("bday") Then
        cmdsave2.Enabled = False
    Else
        cmdsave2.Caption = "SAVE CHANGES"
        cmdsave2.Enabled = True
    End If
End If

End Sub

Private Sub txtEmail2_GotFocus()
frmAddEdit.txtEmail2.SelStart = Len(frmAddEdit.txtEmail2.Text)
End Sub

Private Sub txtEmail2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789-_@., ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtFirst1_Change()
lblmessage1_4.Visible = False
End Sub

Private Sub txtFirst1_GotFocus()
frmAddEdit.txtFirst1.SelStart = Len(frmAddEdit.txtFirst1.Text)
End Sub

Private Sub txtFirst1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ. ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtFirst2_Change()
lblmessage2_1.Visible = False

dtbday = Format(dtbday2, "mmmm dd, yyyy")

If Not cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype") Or Not txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first") Or Not txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle") Or Not txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last") Or Not txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age") Or Not cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex") Or Not txtAddress2.Text = frmMain.accountsdata.Recordset.Fields("address") Or Not txtPhone2.Text = frmMain.accountsdata.Recordset.Fields("phone") Or Not txtEmail2.Text = frmMain.accountsdata.Recordset.Fields("email") Then
    cmdsave2.Enabled = True
    cmdsave2.Caption = "SAVE CHANGES"
Else
    cmdsave2.Enabled = False
    If dtbday = frmMain.accountsdata.Recordset.Fields("bday") Then
        cmdsave2.Enabled = False
    Else
        cmdsave2.Caption = "SAVE CHANGES"
        cmdsave2.Enabled = True
    End If
End If

End Sub

Private Sub txtFirst2_GotFocus()
frmAddEdit.txtFirst2.SelStart = Len(frmAddEdit.txtFirst2.Text)
End Sub

Private Sub txtFirst2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ. ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtLast1_Change()
lblmessage1_6.Visible = False
End Sub

Private Sub txtLast1_GotFocus()
frmAddEdit.txtLast1.SelStart = Len(frmAddEdit.txtLast1.Text)
End Sub

Private Sub txtLast1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ. ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtLast2_Change()
lblmessage2_2.Visible = False
dtbday = Format(dtbday2, "mmmm dd, yyyy")

If Not cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype") Or Not txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first") Or Not txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle") Or Not txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last") Or Not txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age") Or Not cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex") Or Not txtAddress2.Text = frmMain.accountsdata.Recordset.Fields("address") Or Not txtPhone2.Text = frmMain.accountsdata.Recordset.Fields("phone") Or Not txtEmail2.Text = frmMain.accountsdata.Recordset.Fields("email") Then
    cmdsave2.Enabled = True
    cmdsave2.Caption = "SAVE CHANGES"
Else
    cmdsave2.Enabled = False
    If dtbday = frmMain.accountsdata.Recordset.Fields("bday") Then
        cmdsave2.Enabled = False
    Else
        cmdsave2.Caption = "SAVE CHANGES"
        cmdsave2.Enabled = True
    End If
End If

End Sub

Private Sub txtLast2_GotFocus()
frmAddEdit.txtLast2.SelStart = Len(frmAddEdit.txtLast2.Text)
End Sub

Private Sub txtLast2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ. ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtMiddle1_Change()
lblmessage1_5.Visible = False
End Sub

Private Sub txtMiddle1_GotFocus()
frmAddEdit.txtMiddle1.SelStart = Len(frmAddEdit.txtMiddle1.Text)
End Sub

Private Sub txtMiddle1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ. ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtMiddle2_Change()
dtbday = Format(dtbday2, "mmmm dd, yyyy")

If Not cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype") Or Not txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first") Or Not txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle") Or Not txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last") Or Not txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age") Or Not cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex") Or Not txtAddress2.Text = frmMain.accountsdata.Recordset.Fields("address") Or Not txtPhone2.Text = frmMain.accountsdata.Recordset.Fields("phone") Or Not txtEmail2.Text = frmMain.accountsdata.Recordset.Fields("email") Then
    cmdsave2.Enabled = True
    cmdsave2.Caption = "SAVE CHANGES"
Else
    cmdsave2.Enabled = False
    If dtbday = frmMain.accountsdata.Recordset.Fields("bday") Then
        cmdsave2.Enabled = False
    Else
        cmdsave2.Caption = "SAVE CHANGES"
        cmdsave2.Enabled = True
    End If
End If

End Sub

Private Sub txtMiddle2_GotFocus()
frmAddEdit.txtMiddle2.SelStart = Len(frmAddEdit.txtMiddle2.Text)
End Sub

Private Sub txtMiddle2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ. ", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtPassword1_Change()
lblmessage1_2.Visible = False
End Sub

Private Sub txtPassword1_GotFocus()
frmAddEdit.txtPassword1.SelStart = Len(frmAddEdit.txtPassword1.Text)
End Sub

Private Sub txtPassword1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789._", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtPhone1_Change()
lblmessage1_8.Visible = False
End Sub

Private Sub txtPhone1_GotFocus()
frmAddEdit.txtPhone1.SelStart = Len(frmAddEdit.txtPhone1.Text)
End Sub

Private Sub txtPhone1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("0123456789-+", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtPhone2_Change()
lblmessage2_5.Visible = False
dtbday = Format(dtbday2, "mmmm dd, yyyy")

If Not cmbUserType2 = frmMain.accountsdata.Recordset.Fields("usertype") Or Not txtFirst2.Text = frmMain.accountsdata.Recordset.Fields("first") Or Not txtMiddle2.Text = frmMain.accountsdata.Recordset.Fields("middle") Or Not txtLast2.Text = frmMain.accountsdata.Recordset.Fields("last") Or Not txtAge2.Text = frmMain.accountsdata.Recordset.Fields("age") Or Not cmbSex2 = frmMain.accountsdata.Recordset.Fields("sex") Or Not txtAddress2.Text = frmMain.accountsdata.Recordset.Fields("address") Or Not txtPhone2.Text = frmMain.accountsdata.Recordset.Fields("phone") Or Not txtEmail2.Text = frmMain.accountsdata.Recordset.Fields("email") Then
    cmdsave2.Enabled = True
    cmdsave2.Caption = "SAVE CHANGES"
Else
    cmdsave2.Enabled = False
    If dtbday = frmMain.accountsdata.Recordset.Fields("bday") Then
        cmdsave2.Enabled = False
    Else
        cmdsave2.Caption = "SAVE CHANGES"
        cmdsave2.Enabled = True
    End If
End If
End Sub

Private Sub txtPhone2_GotFocus()
frmAddEdit.txtPhone2.SelStart = Len(frmAddEdit.txtPhone2.Text)
End Sub

Private Sub txtPhone2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("0123456789-+", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtretypepassword1_Change()
lblmessage1_3.Visible = False
End Sub

Private Sub txtretypepassword1_GotFocus()
frmAddEdit.txtretypepassword1.SelStart = Len(frmAddEdit.txtretypepassword1.Text)
End Sub

Private Sub txtretypepassword1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789._", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtUsername1_Change()
lblmessage1_1.Visible = False
End Sub

Private Sub txtUsername1_GotFocus()
frmAddEdit.txtUsername1.SelStart = Len(frmAddEdit.txtUsername1.Text)
End Sub

Private Sub txtUsername1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If InStr("abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789._", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
