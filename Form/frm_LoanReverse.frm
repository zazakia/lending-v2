VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_LoanReverse 
   Caption         =   "frmLoanReverse"
   ClientHeight    =   10575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18645
   LinkTopic       =   "Form1"
   ScaleHeight     =   10575
   ScaleWidth      =   18645
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   20520
      TabIndex        =   58
      Top             =   1920
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   450
      _Version        =   393216
      Format          =   530710529
      CurrentDate     =   41894
   End
   Begin VB.Frame Frame1 
      Caption         =   "Loan Reverse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20175
      Begin VB.Frame Frame6 
         Height          =   1935
         Left            =   14760
         TabIndex        =   70
         Top             =   1200
         Width           =   3135
         Begin VB.TextBox txtOverToCus 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   79
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtNotPosted 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   76
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtCharges 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   1320
            TabIndex        =   71
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Over to Cus:"
            Height          =   195
            Left            =   240
            TabIndex        =   78
            Top             =   1560
            Width           =   885
         End
         Begin VB.Label lblLabel23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Not Posted:"
            Height          =   195
            Left            =   240
            TabIndex        =   77
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Charges:"
            Height          =   195
            Left            =   240
            TabIndex        =   72
            Top             =   960
            Width           =   630
         End
      End
      Begin VB.TextBox txtFCollector 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         TabIndex        =   62
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   4680
         TabIndex        =   53
         Top             =   4920
         Width           =   3135
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4335
         Left            =   120
         TabIndex        =   52
         Top             =   5400
         Width           =   19935
         _ExtentX        =   35163
         _ExtentY        =   7646
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13321
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13321
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3720
         TabIndex        =   47
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtDate 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   11160
         TabIndex        =   42
         Top             =   240
         Width           =   3135
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   4680
         Top             =   360
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   2880
         TabIndex        =   41
         Top             =   9840
         Width           =   1815
      End
      Begin VB.CommandButton btnReverse 
         Caption         =   "&Reverse"
         Height          =   615
         Left            =   120
         TabIndex        =   40
         Top             =   9840
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Left            =   12360
         TabIndex        =   33
         Top             =   3120
         Width           =   5655
         Begin VB.Frame Frame5 
            Height          =   135
            Left            =   2400
            TabIndex        =   69
            Top             =   0
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox txtDateP 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   405
            Left            =   2160
            TabIndex        =   56
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox txtLoanTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2160
            TabIndex        =   36
            Top             =   1200
            Width           =   3255
         End
         Begin VB.TextBox txtTotalCharges 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2160
            TabIndex        =   35
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox txtP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2160
            TabIndex        =   34
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label2 
            Caption         =   "Date Loaned:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Total      :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1080
            TabIndex        =   39
            Top             =   1200
            Width           =   945
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Less: Total Charges  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Principal                        :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   1905
         End
         Begin VB.Line Line1 
            X1              =   1920
            X2              =   5400
            Y1              =   1080
            Y2              =   1080
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Loan Breakdown Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6600
         TabIndex        =   12
         Top             =   3120
         Width           =   5655
         Begin VB.TextBox txtTotalAmortization 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2160
            TabIndex        =   31
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Total Amortization     :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   32
            Top             =   480
            Width           =   1875
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Charges Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   8040
         TabIndex        =   11
         Top             =   1080
         Width           =   6615
         Begin VB.TextBox txtPassbook 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4800
            TabIndex        =   48
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtPenalty 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4800
            TabIndex        =   29
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtDelivery 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   4800
            TabIndex        =   28
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtLifeInsurance 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   26
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtServicefee 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   21
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtCollection 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   20
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtFireInsurance 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1320
            TabIndex        =   19
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblPassbook 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Passbook:"
            Height          =   195
            Left            =   3840
            TabIndex        =   49
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Penalty     :"
            Height          =   195
            Left            =   3840
            TabIndex        =   30
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Delivery    :"
            Height          =   195
            Left            =   3840
            TabIndex        =   27
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "BALANCE :"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   825
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Service fee     :"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   1080
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Collection        :"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fire Insurance :"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtdtMaturity 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txtdtRelease 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox txtTotal 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtPrincipal 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txtCustomer 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtCollector 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   265551873
         CurrentDate     =   41870
      End
      Begin VB.Label lblCuCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6600
         TabIndex        =   75
         Top             =   600
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblRWASD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   4680
         TabIndex        =   74
         Top             =   4560
         Width           =   75
      End
      Begin VB.Label lblLabel22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH USING CUSTOMER NAME , CUSTOMER CODE OR DATE"
         Height          =   195
         Left            =   4800
         TabIndex        =   73
         Top             =   4560
         Width           =   4995
      End
      Begin VB.Label lblCollateral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6960
         TabIndex        =   68
         Top             =   960
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblCollectorCharges 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   3240
         TabIndex        =   67
         Top             =   720
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblTotalPayment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6240
         TabIndex        =   66
         Top             =   720
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblLoanDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6360
         TabIndex        =   65
         Top             =   480
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblCollectorCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6000
         TabIndex        =   64
         Top             =   600
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblCollwe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4080
         TabIndex        =   63
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   720
         TabIndex        =   61
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label lblCodeMust 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code must be Typed by Numbers only.Then press ENTER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   60
         Top             =   1560
         Width           =   5265
      End
      Begin VB.Label lblSearchHere 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Here:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3120
         TabIndex        =   54
         Top             =   4920
         Width           =   1410
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6000
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Date     :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10200
         TabIndex        =   43
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Maturity           :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   18
         Top             =   4080
         Width           =   1560
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Date Release :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   17
         Top             =   3720
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total                 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   16
         Top             =   3360
         Width           =   1605
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Principal           :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   15
         Top             =   3000
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Customer  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   14
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Code          :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   13
         Top             =   1920
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Collector    :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   525
      End
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   195
      Left            =   20400
      TabIndex        =   59
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblDateEncoded 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Left            =   20400
      TabIndex        =   55
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   10560
      TabIndex        =   46
      Top             =   4440
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblUserlevel 
      AutoSize        =   -1  'True
      Caption         =   "Label22"
      Height          =   195
      Left            =   12720
      TabIndex        =   51
      Top             =   4920
      Width           =   570
   End
   Begin VB.Label lblTime 
      Caption         =   "Label22"
      Height          =   255
      Left            =   12960
      TabIndex        =   50
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Label22"
      Height          =   195
      Left            =   12840
      TabIndex        =   45
      Top             =   3480
      Width           =   570
   End
   Begin VB.Label lblLoanID 
      Height          =   255
      Left            =   12960
      TabIndex        =   44
      Top             =   4200
      Width           =   255
   End
End
Attribute VB_Name = "frm_LoanReverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub lenth()

        '<EhHeader>
        On Error GoTo lenth_Err

        TxtLog "Entered lenth"

        '</EhHeader>

100     If Len(txtCode.Text) = 1 Then
            'For 1 to 9
102         lblCount.Caption = "0000" + txtCode.Text
104         txtCode.Text = lblCount.Caption
            'for 10 to 99
106     ElseIf Len(txtCode.Text) = 2 Then
108         lblCount.Caption = "000" + txtCode.Text
110         txtCode.Text = lblCount.Caption
            'for 100 to 999
112     ElseIf Len(txtCode.Text) = 3 Then
114         lblCount.Caption = "00" + txtCode.Text
116         txtCode.Text = lblCount.Caption
118     ElseIf Len(txtCode.Text) = 4 Then
120         lblCount.Caption = "0" + txtCode.Text
122         txtCode.Text = lblCount.Caption
        End If

        '<EhFooter>

        TxtLog "Exited lenth"

        Exit Sub

lenth_Err:
        ErrReport Err.Description, "LendingClient.frm_LoanReverse.lenth", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_LoanReverse
'    Project    : Project1
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Sub sort2()

        '<EhHeader>
        On Error GoTo sort2_Err

        TxtLog "Entered sort2"

        '</EhHeader>

100     If rsLoan.State = 1 Then rsLoan.Close
102     rsLoan.Open "Select * from tblLoan Order by LoanID desc"
104     Set DataGrid1.DataSource = rsLoan
        
106     DataGrid1.Width = Me.Width
108     DataGrid1.Columns(1).Width = 1100
110     DataGrid1.Columns(2).Width = 1100
112     DataGrid1.Columns(4).Width = 1500
114     DataGrid1.Columns(5).Width = 950
116     DataGrid1.Columns(6).Width = 750
118     DataGrid1.Columns(7).Width = 1200
120     DataGrid1.Columns(8).Width = 1300
122     DataGrid1.Columns(9).Width = 1100
124     DataGrid1.Columns(10).Width = 1100
126     DataGrid1.Columns(22).Width = 880

        '<EhFooter>

        TxtLog "Exited sort2"

        Exit Sub

sort2_Err:
        ErrReport Err.Description, "LendingClient.frm_LoanReverse.sort2", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnClose_Click()

        '<EhHeader>
        On Error GoTo btnClose_Click_Err

        TxtLog "Entered btnClose_Click"

        '</EhHeader>

100     If btnClose.Caption = "&Close" Then
102         Unload Me
        Else
104         txtSearch.Enabled = True
106         btnClose.Caption = "&Close"
108         btnReverse.Caption = "&Reverse"
110         txtCode.Text = ""
112         txtCollector.Text = ""
114         txtFCollector.Text = ""
116         txtPrincipal.Text = ""
118         txtTotal.Text = ""
120         txtServicefee.Text = "0"
122         txtLifeInsurance.Text = ""
124         txtdtMaturity.Text = ""
126         txtdtRelease.Text = ""
128         txtTotalAmortization.Text = ""
130         txtCustomer.Text = ""
132         txtFirstname.Text = ""
134         txtFireInsurance.Text = ""
136         txtCollection.Text = "0"
138         txtServicefee.Text = "0"
140         txtDelivery.Text = "0"
142         txtPenalty.Text = ""
144         txtPassbook.Text = ""
146         txtCode.Enabled = False
            '  txtCharge.Text = "0"
148         Call sort2
            ' txtDateRelease.Enabled = False
150         Call Timer1_Timer
            
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_LoanReverse.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnReverse_Click()

        '<EhHeader>
        On Error GoTo btnReverse_Click_Err

        TxtLog "Entered btnReverse_Click"

        '</EhHeader>

100     txtSearch.Enabled = False

102     If btnReverse.Caption = "&Reverse" Then
104         btnReverse.Caption = "&Update"
106         btnClose.Caption = "&Cancel"
108         txtCode.Enabled = True
110         txtCode.SetFocus
    
112     ElseIf btnReverse.Caption = "&Update" Then

114         If txtCustomer.Text = "" Or txtCollector.Text = "" Then
116             MsgBox "No Record Found", vbInformation, "Webplus Lending Corporation"
118             txtCode.SetFocus
            Else

                'say gamit ani na code?
120             If rsCustomer.State = 1 Then rsCustomer.Close
122             rsCustomer.Open "Select * from tblCustomer where Code = '" & _
                        lblCuCode.Caption & "'" 'get customer of the loan
                'rsCustomer!Balance = Val(txtTotalAmortization.Text) - Val(txtTotal.Text)
                'rsCustomer!Balance = Val(txtLifeInsurance.Text)
                'rsCustomer.Update

                Dim RLoan As String

124             RLoan = ""
126             RLoan = lblCuCode.Caption
                
                'dili ma reverse ang full paid na loans?
128             If rsLoan.State = 1 Then rsLoan.Close
130             rsLoan.Open "Select * from tblLoan where Code = '" & RLoan & _
                        "' and (Status =  'Full Paid' or Status = 'Good') Order By LoanID desc"

132             If rsLoan.RecordCount = 0 Then
134                 MsgBox "No Record Found", vbInformation, "Webplus Lending Corporation"
                Else

136                 If MsgBox("Are you sure you want to Reverse this Transaction?", _
                            vbQuestion + vbYesNo, "Webplus Lending Corporation") = _
                            vbYes Then

138                     With rsLoan
                            'changes the existing loan status to reversed
140                         !Status = "Reversed"
142                         !LoanStatus = "Reversed"
144                         .AddNew
146                         !Collector = txtCollector.Text
148                         !CollectorFname = txtFCollector.Text
150                         !CollectorCode = lblCollectorCode.Caption
152                         !code = RLoan
154                         !Customer = txtCustomer.Text
156                         !firstname = txtFirstname.Text
158                         !principal = txtPrincipal.Text
160                         !total = txtTotal.Text
162                         !DateRelease = txtdtRelease.Text
164                         !Maturity = txtdtMaturity.Text
166                         !FireInsurance = txtFireInsurance.Text
168                         !delivery = txtDelivery.Text
170                         !collection = txtCollection.Text
172                         !Servicefee = txtServicefee.Text
174                         !Balance = txtLifeInsurance.Text
176                         !Status = "Reversing"
178                         !Penalty = txtPenalty.Text
180                         !Passbook = txtPassbook.Text
182                         !TotalAmortization = "0"
184                         !TotalCharges = "0"
186                         !LoanTotal = "0"
188                         !LoanStatus = "Reversing"
190                         !User = lblUser.Caption
192                         !LoanDate = lblLoanDate.Caption
194                         !TotalPayment = lblTotalPayment.Caption
196                         !NotPosted = Val(txtNotPosted.Text)
198                         !CollectorCharge = Val(lblCollectorCharges.Caption)
200                         !CollectorCharge = Val(txtCharges.Text)
202                         !Collateral = lblCollateral.Caption
204                         .Update
                        End With
                        
                        'say gamit ani na code?
206                     If rsLoan.State = 1 Then rsLoan.Close
                        'pfft.
                        'rsLoan.Open "Select * from tblLoan where Code = '" & RLoan & "' and Status =  '" & "Full Paid" & "' and TotalAmortization <> 0"
208                     rsLoan.Open "Select * from tblLoan where Code = '" & RLoan & _
                                "' and (Status =  'Full Paid' or Status = 'Good') Order By LoanID desc"
                                               
                        'If rsLoan.State = 1 Then rsLoan.Close
                        'rsLoan.Open "Select * from "
                        
                        'diri kitaon niya kung naa ba siyay laing loan besides sa loan na gi reverse
210                     If rsLoan.RecordCount <> 0 Then 'kung naa
212                         rsCustomer!Balance = rsLoan!TotalAmortization
                            
214                         If rsLoan!TotalAmortization = 0 Then

216                             With rsLoan
218                                 !Status = "Full Paid"
220                                 .Update
                                End With
                             
222                         ElseIf rsLoan!TotalAmortization > 0 Then

224                             With rsLoan
226                                 !Status = "Good"
228                                 .Update
                                End With

                            Else
230                             MsgBox ( _
                                        "Please check the current loan for it has a negative balance. Thank you.")
                            End If
                            
                            'rsCustomer.Update
                        Else 'kung wala
232                         rsCustomer!Balance = 0
                        End If
                        
234                     rsCustomer.Update
                        
                        'what -start
                        'If rsLoan.RecordCount <> 0 Then
                        'rsLoan.MoveLast

                        'Updates the status
                        'With rsLoan
                        '!Status = "Good"
                        '.Update
                        'End With

                        'Else
                            
                        'End If
                        'what -end

236                     With rsTrail
238                         .AddNew
240                         !UserName = lblUser.Caption
242                         !userlevel = lblUserlevel.Caption
244                         !Activity = "Reversing Loan"
246                         !Time = lblTime.Caption
248                         !Date = txtDate.Text
250                         .Update
                        End With

252                     MsgBox "Record Successfully Reversed", vbInformation, _
                                "Webplus Lending Corporation"
254                     Unload Me
256                     frm_LoanReverse.lblUser.Caption = MDIForm1.lblUserName.Caption
258                     Me.Show
260                     txtSearch.Enabled = True
                    End If
                End If
            End If
        End If

        'End If

        '<EhFooter>

        TxtLog "Exited btnReverse_Click"

        Exit Sub

btnReverse_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_LoanReverse.btnReverse_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DataGrid1_KeyPress_Err

        TxtLog "Entered DataGrid1_KeyPress"

        '</EhHeader>

100     KeyAscii = 0

        '<EhFooter>

        TxtLog "Exited DataGrid1_KeyPress"

        Exit Sub

DataGrid1_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_LoanReverse.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call Loan
104     Call Loan1
106     Call Customer
108     Call payment

110     If rsLoan.State = 1 Then rsLoan.Close
        'Sort records to descending
112     rsLoan.Open "Select * from tblLoan Order By LoanID desc"
        'Display records on Datagrid
114     Set DataGrid1.DataSource = rsLoan
        
116     DataGrid1.Width = Me.Width
118     DataGrid1.Columns(1).Width = 1100
120     DataGrid1.Columns(2).Width = 1100
122     DataGrid1.Columns(4).Width = 1500
124     DataGrid1.Columns(5).Width = 950
126     DataGrid1.Columns(6).Width = 750
128     DataGrid1.Columns(7).Width = 1200
130     DataGrid1.Columns(8).Width = 1300
132     DataGrid1.Columns(9).Width = 1100
134     DataGrid1.Columns(10).Width = 1100
136     DataGrid1.Columns(22).Width = 880

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_LoanReverse.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    txtDate.Text = Date
    lblTime.Caption = Time

End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtCode_KeyPress_Err

        TxtLog "Entered txtCode_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf KeyAscii = 13 Then
            'calls the function that will search the code that don't need to input "0" on code
110         Call lenth

112         If txtCode.Text = "" Then
114             MsgBox "Code should not be blank.", vbInformation, _
                        "Webplus Lending Corporation"
116             txtCode.SetFocus
            Else

118             If rsLoan.State = 1 Then rsLoan.Close
120             rsLoan.Open "Select * from tblLoan where Code = '" & txtCode.Text & _
                        "' and Status <> '" & "Full Paid" & "' and Status <> '" & _
                        "Reversed" & "'and Status <> '" & "Reversing" & "'"

122             If rsLoan.RecordCount = 0 Then
124                 MsgBox "No Record found.", vbInformation, _
                            "Webplus Lending Corporation."
126                 txtCode.Text = ""
128                 txtCode.SetFocus
                Else

130                 If rsPayment.State = 1 Then rsPayment.Close
132                 rsPayment.Open "Select * from tblPayment where LoanID = " & _
                            rsLoan!LoanID & " and Status = '" & "Good" & "'"

134                 If rsPayment.RecordCount <> 0 Then
136                     MsgBox _
                                "This Loan as already have a payment. Reverse the existing payment(s) first if you wish to reverse this loan. ", _
                                vbInformation, "Webplus Lending Corporation"
138                     txtCode.SetFocus
                    Else

                        'If record found , It will show the customer record based on customer loan
140                     With rsLoan
142                         lblLoanID.Caption = !LoanID
144                         txtCollector.Text = !Collector
146                         txtFCollector.Text = !CollectorFname
148                         txtCustomer.Text = !Customer
150                         txtFirstname.Text = !firstname
152                         txtPrincipal.Text = !principal
154                         txtTotal.Text = !total
156                         txtdtRelease.Text = !DateRelease
158                         txtdtMaturity.Text = !Maturity
160                         txtFireInsurance.Text = !FireInsurance
162                         txtCollection.Text = !collection
164                         txtServicefee.Text = !Servicefee
166                         txtLifeInsurance.Text = !Balance
168                         txtDelivery.Text = !delivery
170                         txtPenalty.Text = !Penalty
172                         txtPassbook.Text = !Passbook
174                         txtTotalAmortization.Text = !TotalAmortization
176                         txtTotalCharges.Text = !TotalCharges
178                         txtLoanTotal.Text = !LoanTotal
180                         txtP.Text = !principal
182                         DTPicker2.Value = !DateRelease
184                         txtDateP.Text = !LoanDate
186                         txtNotPosted.Text = !NotPosted
188                         lblCuCode.Caption = !code
190                         lblTotal.Caption = Val(txtLoanTotal.Text)
192                         lblDateEncoded.Caption = !LoanDate
194                         lblCollectorCode.Caption = !CollectorCode
196                         lblLoanDate.Caption = !LoanDate
198                         lblTotalPayment.Caption = !TotalPayment
200                         txtCharges.Text = !CollectorCharge
202                         lblCollectorCharges.Caption = !CollectorCharge

                            'not posted
                            'over to cus
204                         If IsNull(!NotPosted) Then
206                             txtNotPosted = 0
                            Else
208                             txtNotPosted = !NotPosted
                            End If
                            
210                         If IsNull(!OverToCus) Then
212                             txtOverToCus = 0
                            Else
214                             txtOverToCus = !OverToCus
                            End If
                            
                        End With
                
                    End If
                End If
            End If

        Else
216         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtCode_KeyPress"

        Exit Sub

txtCode_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_LoanReverse.txtCode_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsLoan.State = 1 Then rsLoan.Close
102     rsLoan.Open "Select * from tblLoan where Customer like '" & txtSearch.Text & _
                "%' or Code like '" & txtSearch.Text & "%' or LoanID like '" & _
                txtSearch.Text & "%' Order by LoanID Desc "
        
104     Set DataGrid1.DataSource = rsLoan

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_LoanReverse.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

