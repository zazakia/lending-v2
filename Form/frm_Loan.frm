VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Loan 
   Caption         =   "frmLoan"
   ClientHeight    =   9360
   ClientLeft      =   1425
   ClientTop       =   405
   ClientWidth     =   19200
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   19200
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "REGULAR LOAN FILE  MAINTENANCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20055
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   61
         Top             =   6120
         Width           =   19695
         _ExtentX        =   34740
         _ExtentY        =   3625
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
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
               Type            =   1
               Format          =   """ ""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13321
               SubFormatType   =   2
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
      Begin VB.TextBox txtLifeInsurance 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   9000
         TabIndex        =   99
         Top             =   4920
         Width           =   3975
      End
      Begin VB.TextBox txtMature 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   8040
         TabIndex        =   84
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Frame Period 
         Caption         =   "Period"
         Height          =   1215
         Left            =   240
         TabIndex        =   75
         Top             =   3960
         Width           =   4815
         Begin VB.OptionButton OptMonth 
            Caption         =   "1.5 Months"
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   80
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton OptMonth 
            Caption         =   "1 Month"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   79
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptMonth 
            Caption         =   "2 Months"
            Height          =   495
            Index           =   2
            Left            =   1440
            TabIndex        =   78
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptMonth 
            Caption         =   "3 Months"
            Height          =   495
            Index           =   4
            Left            =   2760
            TabIndex        =   77
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptMonth 
            Caption         =   "2.5 Months"
            Height          =   495
            Index           =   3
            Left            =   1440
            TabIndex        =   76
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "LOAN TYPE"
         Height          =   855
         Left            =   240
         TabIndex        =   68
         Top             =   3000
         Width           =   3135
         Begin VB.OptionButton OptType 
            Caption         =   "Emergency"
            Height          =   495
            Index           =   1
            Left            =   1560
            TabIndex        =   70
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton OptType 
            Caption         =   "Regular"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox txtTerms 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5520
         TabIndex        =   67
         Text            =   "30"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtInterestRate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   65
         Text            =   "6"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   16080
         TabIndex        =   13
         Top             =   840
         Width           =   2415
         Begin VB.TextBox txtOverToCus 
            Height          =   405
            Left            =   240
            TabIndex        =   29
            Text            =   "0"
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblAddedTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Added to Total Charges"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   89
            Top             =   360
            Width           =   1920
         End
         Begin VB.Label Label10 
            Caption         =   "Over to Customer:"
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
            TabIndex        =   28
            Top             =   0
            Width           =   2175
         End
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   18720
         TabIndex        =   58
         Top             =   2640
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         _Version        =   393216
         Format          =   119209985
         CurrentDate     =   41939
      End
      Begin VB.TextBox txtFCollector 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   10560
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.Frame lll 
         Caption         =   "Collector Charges"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   13320
         TabIndex        =   12
         Top             =   840
         Width           =   2535
         Begin VB.TextBox txtNotPosted 
            Height          =   375
            Left            =   1200
            TabIndex        =   34
            Text            =   "0"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtCharge 
            Height          =   405
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   27
            Text            =   "0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblDeductedTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Deducted to Total Charges"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   360
            TabIndex        =   88
            Top             =   360
            Width           =   1920
         End
         Begin VB.Label lblLabel10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Not Posted:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   360
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblCharge 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Charge :"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   720
            Width           =   600
         End
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1920
         TabIndex        =   22
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   8040
         TabIndex        =   32
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtCollector 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   8040
         TabIndex        =   9
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtFirstname 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   10560
         TabIndex        =   33
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtLoanID 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   18000
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   2040
         TabIndex        =   60
         Top             =   5640
         Width           =   5055
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   3480
         TabIndex        =   64
         Top             =   8280
         Width           =   1815
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         Height          =   615
         Left            =   720
         TabIndex        =   63
         Top             =   8280
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   13320
         TabIndex        =   38
         Top             =   2760
         Width           =   5295
         Begin VB.TextBox txtLoanTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2880
            TabIndex        =   53
            Text            =   "txtLoanTotal"
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox txtTotalCharges 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2880
            TabIndex        =   44
            Text            =   "txtTotalCharges"
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2880
            TabIndex        =   40
            Text            =   "txtP"
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label lblTxtLoanTotalAs 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "txtLoanTotal as Total Loan Released"
            Height          =   675
            Left            =   120
            TabIndex        =   98
            Top             =   1440
            Visible         =   0   'False
            Width           =   2625
         End
         Begin VB.Line Line2 
            BorderColor     =   &H008080FF&
            BorderWidth     =   7
            X1              =   2760
            X2              =   5160
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label22 
            Caption         =   "Total for Release:"
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
            TabIndex        =   52
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label21 
            Caption         =   "Less: Total Charges"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Principal                      :"
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
            TabIndex        =   39
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Loan Breakdown Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5640
         TabIndex        =   45
         Top             =   1560
         Width           =   4695
         Begin VB.TextBox txtTotalAmortization 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   2400
            TabIndex        =   49
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Total Amortization  :"
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
            TabIndex        =   51
            Top             =   480
            Width           =   1740
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Charges Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   7440
         TabIndex        =   11
         Top             =   3000
         Width           =   5655
         Begin VB.TextBox txtPassbook 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   3960
            MaxLength       =   3
            TabIndex        =   95
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtServicefee 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   1560
            TabIndex        =   94
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtDelivery 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3960
            MaxLength       =   3
            TabIndex        =   93
            Top             =   360
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtCollection 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3960
            MaxLength       =   4
            TabIndex        =   92
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtPenalty 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   25
            Top             =   885
            Width           =   1455
         End
         Begin VB.TextBox txtFireInsurance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   18
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Fee    :"
            Height          =   195
            Left            =   240
            TabIndex        =   97
            Top             =   1440
            Width           =   1080
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Passbook         :"
            Height          =   195
            Left            =   3120
            TabIndex        =   96
            Top             =   1440
            Width           =   1155
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Penalty             :"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label14 
            Caption         =   "Collection             :"
            Height          =   255
            Left            =   3120
            TabIndex        =   23
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Delivery            :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3120
            TabIndex        =   19
            Top             =   480
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Insurance   :"
            Height          =   315
            Left            =   240
            TabIndex        =   17
            Top             =   480
            Width           =   1125
         End
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   16200
         TabIndex        =   41
         Text            =   "txtTotal"
         Top             =   4560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtPrincipal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   37
         Top             =   1620
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   18720
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   264896513
         CurrentDate     =   41866
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   4680
         Top             =   480
      End
      Begin MSComCtl2.DTPicker dtDateRelease 
         Height          =   615
         Left            =   1920
         TabIndex        =   85
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   264896513
         CurrentDate     =   41870
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "OLD BALANCE  :"
         Height          =   195
         Left            =   7680
         TabIndex        =   100
         Top             =   5040
         Width           =   1245
      End
      Begin VB.Label lblCodeMust 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Required Field"
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
         Left            =   600
         TabIndex        =   91
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   360
         TabIndex        =   90
         Top             =   720
         Width           =   180
      End
      Begin VB.Line Line3 
         BorderStyle     =   4  'Dash-Dot
         X1              =   5280
         X2              =   5280
         Y1              =   240
         Y2              =   5160
      End
      Begin VB.Line Line1 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   240
         X2              =   18840
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Shape Shape1 
         Height          =   1095
         Left            =   10560
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   13560
         TabIndex        =   87
         Top             =   4680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   5640
         TabIndex        =   86
         Top             =   2520
         Width           =   1560
      End
      Begin VB.Label lblDays 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "days"
         Height          =   195
         Left            =   6960
         TabIndex        =   83
         Top             =   3480
         Width           =   450
      End
      Begin VB.Label lblLoanID2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan ID"
         Height          =   195
         Left            =   16560
         TabIndex        =   81
         Top             =   0
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interest Rate %"
         Height          =   195
         Left            =   3840
         TabIndex        =   74
         Top             =   3120
         Width           =   1080
      End
      Begin VB.Label lblTerms 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terms"
         Height          =   195
         Left            =   5970
         TabIndex        =   73
         Top             =   3120
         Width           =   435
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payment / Day"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   10800
         TabIndex        =   72
         Top             =   1800
         Width           =   1740
      End
      Begin VB.Label lblPerDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblPerDay"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10920
         TabIndex        =   71
         Top             =   2160
         Width           =   1725
      End
      Begin VB.Label lblInterestRate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Interest Rate"
         Height          =   195
         Left            =   18840
         TabIndex        =   66
         Top             =   3240
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblLifeInsurance 
         Caption         =   "LifeInsurance"
         Height          =   375
         Left            =   18720
         TabIndex        =   48
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblDtrelease 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5040
         TabIndex        =   54
         Top             =   3960
         Width           =   45
      End
      Begin VB.Label lblCuCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5640
         TabIndex        =   50
         Top             =   3840
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   795
         Left            =   1680
         TabIndex        =   56
         Top             =   1800
         Width           =   180
      End
      Begin VB.Label EE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH USING LOAN ID , CUSTOMER CODE OR CUSTOMER NAME"
         Height          =   195
         Left            =   7320
         TabIndex        =   57
         Top             =   5760
         Width           =   5205
      End
      Begin VB.Label lblCollectorFname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblCollectorFname"
         Height          =   195
         Left            =   18600
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblCollectorCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CollectorCode"
         Height          =   195
         Left            =   18840
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblCount"
         Height          =   195
         Left            =   18960
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code must be numbers only.Then press ENTER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   4305
      End
      Begin VB.Label lblSearchHere 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Here :"
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
         TabIndex        =   59
         Top             =   5640
         Width           =   1470
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   1680
         TabIndex        =   15
         Top             =   1200
         Width           =   180
      End
      Begin VB.Label Label8 
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
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Customer    :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   31
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Code            :"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1110
         Width           =   1350
      End
      Begin VB.Label Label3 
         Caption         =   "Collector      :"
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
         Left            =   5640
         TabIndex        =   8
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lbldate 
         AutoSize        =   -1  'True
         Caption         =   "lbldate"
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
         Left            =   8040
         TabIndex        =   4
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date Today:"
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
         Left            =   5640
         TabIndex        =   3
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.Label lblClickAdd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click add button to ad new Loan"
      Height          =   195
      Left            =   360
      TabIndex        =   82
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label lblBalance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   435
      Left            =   10320
      TabIndex        =   62
      Top             =   120
      Width           =   930
   End
   Begin VB.Label lblUser 
      Caption         =   "Label18"
      Height          =   255
      Left            =   4560
      TabIndex        =   55
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblUserlevel 
      AutoSize        =   -1  'True
      Caption         =   "Label18"
      Height          =   195
      Left            =   6240
      TabIndex        =   42
      Top             =   240
      Width           =   3930
   End
   Begin VB.Label lblTime 
      Height          =   495
      Left            =   11520
      TabIndex        =   20
      Top             =   1320
      Width           =   135
   End
   Begin VB.Label lblLoanID 
      Height          =   375
      Left            =   12840
      TabIndex        =   47
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label lblPrincipal 
      Caption         =   "Label23"
      Height          =   255
      Left            =   3360
      TabIndex        =   30
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frm_Loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Loan
'    Project    : LendingClient
'
'    Description: Lending Business Management System with Payroll
'
'    Modified   : Brayan Lee A. Bautista
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       lenth
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Sub lenth()

        '<EhHeader>
        On Error GoTo lenth_Err

        TxtLog "Entered lenth"

        '</EhHeader>

100     If Len(txtCode.Text) = 1 Then
102         lblCount.Caption = "0000" + txtCode.Text
104         txtCode.Text = lblCount.Caption
106     ElseIf Len(txtCode.Text) = 2 Then
108         lblCount.Caption = "000" + txtCode.Text
110         txtCode.Text = lblCount.Caption
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
        ErrReport Err.Description, "LendingClient.frm_Loan.lenth", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       sort
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Sub sort()

        '<EhHeader>
        On Error GoTo sort_Err

        TxtLog "Entered sort"

        '</EhHeader>

100     If rsLoan.State = 1 Then rsLoan.Close
        'Sort records descending
102     rsLoan.Open "Select * from tblLoan  Order By LoanID desc"
104     Set DataGrid1.DataSource = rsLoan
106     DataGrid1.Width = Me.Width
136     DataGrid1.Columns(5).Caption = "Encoded Date"
108     DataGrid1.Columns(1).Width = 1100
110     DataGrid1.Columns(2).Width = 1100
112     DataGrid1.Columns(4).Width = 1500
114     DataGrid1.Columns(5).Width = 950
116     DataGrid1.Columns(6).Width = 750
118     DataGrid1.Columns(7).Width = 1000
120     DataGrid1.Columns(8).Width = 1300
122     DataGrid1.Columns(9).Width = 1100
124     DataGrid1.Columns(10).Width = 1100
126     DataGrid1.Columns(11).Width = 1100
128     DataGrid1.Columns(22).Width = 880

        '<EhFooter>

        TxtLog "Exited sort"

        Exit Sub

sort_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.sort", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       SumCharges
' Description:       for computation in loan form all txt box totals
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Sub SumCharges()

        '<EhHeader>
        On Error GoTo SumCharges_Err

        TxtLog "Entered SumCharges"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

        'This procedure Automatically computes how much the total Charges
        
        Dim sum        As Double

        Dim Servicefee As Double

        'Computing the processing fee , depending how much the Principal or Loan
        'jajavi pa ni processing fee
        'divideByFiveHundred = Val(txtPrincipal.Text) / 500
        'MAke it editable no computation for renewal fee
        'divideByFiveHundred = Val(txtPrincipal.Text) * 0.01 + 20
        
102     Servicefee = Val(txtServicefee)
        
        'diri na part. dapat i recheck
        'Balance textbox
104     txtLifeInsurance.Text = Val(lblLifeInsurance.Caption) - (Val(txtCharge.Text) + _
                Val(txtNotPosted.Text)) + Val(txtOverToCus.Text)
        
        'can be disabled by removing the code in the if else.
        'recheck code above
        
106     txtTotalAmortization = Val(txtTotal.Text)
108     txtServicefee.Text = Servicefee
110     txtTotalAmortization = Val(txtTotalAmortization.Text) - Val(txtCollection.Text)

        'Computing the Sum of all charges
112     If OptType(1).Value = True Then
114         sum = 0 'EMERGENCY LOAN
        Else
116         sum = Val(txtFireInsurance.Text) + Val(txtCollection.Text) + Val( _
                    txtServicefee.Text) + Val(txtLifeInsurance.Text) + Val( _
                    txtDelivery.Text) + Val(txtPenalty.Text) + Val(txtPassbook.Text)
        End If
        
118     txtTotalCharges.Text = sum
        
        Dim totalsum As Double

120     totalsum = Val(txtP.Text) - Val(txtTotalCharges.Text)
122     txtLoanTotal.Text = Val(totalsum)

        Dim LessSunday As Integer
        
124     Select Case txtTerms
        
            Case 30
126             LessSunday = Val(txtTerms) - 4

128         Case 45
130             LessSunday = Val(txtTerms) - 6

132         Case 60
134             LessSunday = Val(txtTerms) - 8

136         Case 75
138             LessSunday = Val(txtTerms) - 10

140         Case 90
142             LessSunday = Val(txtTerms) - 12
        
        End Select

144     If LessSunday = 0 Then

            Exit Sub

        End If
        
146     If OptType(1) = True Then
148         lblPerDay = Round(Val(txtTotalAmortization) - Val(txtPrincipal), 0)
        Else
150         lblPerDay = Round(txtTotalAmortization / LessSunday, 0)
        End If

        '<EhFooter>

        TxtLog "Exited SumCharges"

        Exit Sub

SumCharges_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.SumCharges", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       btnAdd_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub btnAdd_Click()

        '<EhHeader>
        On Error GoTo btnAdd_Click_Err

        TxtLog "Entered btnAdd_Click"

        '</EhHeader>

100     txtSearch.Enabled = False

        'add new loan not saving
102     If btnAdd.Caption = "&Add" Then
            
104         btnAdd.Caption = "&Save"
106         txtCode.Enabled = True
    
108         txtPrincipal.Enabled = True
110         dtDateRelease.Enabled = True
    
112         txtFireInsurance.Enabled = True
114         txtFireInsurance.Text = "0"
116         txtCollection.Enabled = True
            
118         txtPenalty.Enabled = True
120         txtDelivery.Enabled = False
    
122         DTPicker1.Enabled = True
124         txtCode.SetFocus
126         btnClose.Caption = "&Cancel"
txtCharge.Text = "0"
txtNotPosted.Text = "0"
txtOverToCus.Text = "0"
txtTotalAmortization.Text = "0"
txtP.Text = "0"
txtTotalCharges.Text = "0"
txtLoanTotal.Text = "0"
txtTotal.Text = "0"
lblPerDay.Caption = "0"
txtTerms.Text = "0"
txtInterestRate.Text = "0"

    
        Else ' saving loan

            'add also this code to code textbox to check immedietly if the customer record is reversed.
            'Recheck Computation
            'Mao ne ang mo check if ang customer gereverse na ba siya..
            'Kung gereverse na gane dle na siya makaloan ug balik
            'Pero ug wala pa ge reverse makaloan pa siya
128         If rsCustomer.State = 1 Then rsCustomer.Close
130         rsCustomer.Open "Select * from tblCustomer where Code = " & Val( _
                    txtCode.Text) & " and status = 'Reversed'"

132         If rsCustomer.RecordCount > 0 Then
134             MsgBox "Customer is already reversed. He/she cannot loan again."
            Else
                'if not reversed compute fields
136             Call SumCharges

138             If rsLoan.State = 1 Then rsLoan.Close
                
                'check if there is a loan same date, code and status
140             rsLoan.Open "Select * from tblLoan where DateRelease =  #" & _
                        dtDateRelease.Value & "# and Code = " & Val(lblCuCode.Caption) _
                        & " and Status = '" & "Good" & "'"

142             If rsLoan.RecordCount <> 0 Then
144                 MsgBox "This Customer already has a loan on " & rsLoan!DateRelease
146                 rsLoan.Close
148                 txtCode.SetFocus
150                 Call Timer1_Timer
                Else
     
                    'If Customer Code not provided
152                 If txtCode.Text = "" Then
154                     MsgBox "Code field is blank", vbInformation, _
                                "Webplus Lending Corporation"
156                     txtCode.SetFocus
                        'If any fields is empty
158                 ElseIf txtCollector.Text = "" Then
160                     MsgBox ( _
                                "There is still no collector assigned in that area during the date")
162                 ElseIf txtCustomer.Text = "" Then
164                     MsgBox "No Record Found", vbInformation, _
                                "Webplus Lending Corporation"
166                     txtCode.Text = ""
168                     txtCode.SetFocus
                        
                        'check if nanubra ra ang total sa collector charges og not posted compara sa past balance
170                 ElseIf txtLifeInsurance < 0 Then
172                     MsgBox _
                                "Please check, inputed values. Specially the collector not posted and charges."
                        'If Loaned Amount is greater than the Total Deductions, it won't proceed the transaction
174                 ElseIf Val(txtTotalCharges.Text) > Val(txtPrincipal.Text) Then
176                     MsgBox "Principal amount is less than the total charges.", _
                                vbInformation, "Webplus Lending Corporation"
178                     txtPrincipal.SetFocus
           
                    Else 'Final Save

                        'Check if required fields are not blank
180                     Select Case txtPrincipal

                            Case ""
182                             MsgBox ("Please Fill-up required fields")

                                Exit Sub

                        End Select

                        'Notify user if there are other changes before saving
184                     If MsgBox( _
                                "Are you sure you want to create a loan for this customer?", _
                                vbQuestion + vbYesNo, "Webplus Lending Corporation") = _
                                vbYes Then

                            'If yes
186                         If rsLoan.State = 1 Then rsLoan.Close
188                         rsLoan.Open "Select * from tblLoan where DateRelease =  #" _
                                    & dtDateRelease.Value & "# and Code = " & _
                                    lblCuCode.Caption & " and Status = '" & "Good" & "'"

190                         If rsLoan.RecordCount <> 0 Then
192                             MsgBox "This Customer is already loaned on " & _
                                        rsLoan!DateRelease
194                             rsLoan.Close
196                             txtCode.SetFocus
198                             Call Timer1_Timer
                            Else

200                             If rsLoan.State = 1 Then rsLoan.Close
202                             rsLoan.Open "Select * from tblLoan where Code = " & _
                                        lblCuCode.Caption & " and Status = '" & "Good" _
                                        & "' "
                        
204                             If rsLoan1.State = 1 Then rsLoan1.Close
206                             rsLoan1.Open " Select * from tblLoan where Code = " & _
                                        lblCuCode.Caption & _
                                        " and Status = 'Good' Order By LoanID Desc"
                        
208                             If rsLoan1.RecordCount <> 0 Then

                                    'Update as Fully Paid if the balance is not 0 only if its regular loan. If emergency loan then add a new loan
                                            
210                                 With rsLoan 'fk it.
                            
                                        'check if emergency loan type then dont update the regular loan data of any field
                                        Select Case OptType(0)

                                            Case True ' Regular Loan
212                                             rsLoan1.MoveFirst 'move to the loan with the highest loanID.
214                                             rsLoan1.Resync 'so mo resync jud sia. if ever na update ang selection before ma pa nimo ma update. i resync? ahahahah
216                                             rsLoan1!Status = "Full Paid"
218                                             rsLoan1.Update 'blah. ahhhhh. fk.
220                                             rsLoan1.Close 'ang ako gipang edit sa network wala ni epek

                                            Case False ' Emergency

                                        End Select

                                        Dim cCode As String

222                                     cCode = ""

                                        'This should include the collector's first name, last name, code
224                                     With rsLoan
226                                         .AddNew
228                                         !Collector = txtCollector.Text
230                                         !code = lblCuCode.Caption

232                                         cCode = lblCuCode.Caption
234                                         !Customer = txtCustomer.Text
236                                         !firstname = txtFirstName.Text
238                                         !principal = txtPrincipal.Text
240                                         !total = txtTotal.Text
242                                         !LoanDate = lblDate.Caption
244                                         !DateRelease = dtDateRelease.Value

248                                         !Status = "Good"
250                                         !FireInsurance = Val(txtFireInsurance.Text)
252                                         !CollectorCharge = Val(txtCharge.Text)
254                                         !delivery = Val(txtDelivery.Text)
256                                         !collection = Val(txtCollection.Text)
258                                         !Servicefee = Val(txtServicefee.Text)
                                            'TODO: testing lang

260                                         Select Case OptType(0)

                                                Case True ' Regular Loan
262                                                 !Balance = Val( _
                                                            txtLifeInsurance.Text) ' balance previous loan
246                                         !Maturity = txtMature.Text
264                                             Case False ' Emergency

                                            End Select

266                                         !Penalty = Val(txtPenalty.Text)
268                                         !Passbook = Val(txtPassbook.Text)
270                                         !TotalAmortization = Val( _
                                                    txtTotalAmortization.Text)
272                                         !TotalCharges = Val(txtTotalCharges.Text)
274                                         !LoanTotal = Val(txtLoanTotal.Text) ' total released
276                                         !CollectorCode = lblCollectorCode.Caption
278                                         !CollectorFname = lblCollectorFname.Caption
280                                         !User = lblUser.Caption
282                                         !LoanStatus = "Good"
284                                         !TotalPayment = Val(txtCollection.Text)
286                                         !NotPosted = Val(txtNotPosted.Text)
288                                         !OverToCus = Val(txtOverToCus.Text)

                                            Dim LoanTypeVar As Integer

290                                         Select Case OptType(0)

                                                Case True ' Regular Loan
292                                                 LoanTypeVar = 1 '"Regular"

294                                             Case False ' Emergency
296                                                 LoanTypeVar = 2 '"Emergency"
                                            End Select

298                                         !LOANTYPE = LoanTypeVar
300                                         !LoanPeriod = Val(txtTerms)
302                                         !InterestRate = Val(txtInterestRate)
304                                         !PaymentPerDay = Val(lblPerDay.Caption)
306                                         .Update
                                        End With
                                
                                        'update this to customer balanace
                        
                                    End With

                                Else
                    
308                                 With rsLoan
310                                     .AddNew
                                        'This should include the collector's first name, last name, code
312                                     !Collector = txtCollector.Text
314                                     !code = lblCuCode.Caption
316                                     cCode = lblCuCode.Caption
318                                     !Customer = txtCustomer.Text
320                                     !firstname = txtFirstName.Text
322                                     !principal = txtPrincipal.Text
324                                     !total = txtTotal.Text
326                                     !LoanDate = lblDate.Caption
328                                     !DateRelease = dtDateRelease.Value
330                                     !Maturity = txtMature.Text
332                                     !CollectorCharge = Val(txtCharge.Text)
334                                     !Status = "Good"
336                                     !FireInsurance = Val(txtFireInsurance.Text)
338                                     !delivery = Val(txtDelivery.Text)
340                                     !collection = Val(txtCollection.Text)
342                                     !Servicefee = Val(txtServicefee.Text)
344                                     !Balance = Val(txtLifeInsurance.Text)
346                                     !Penalty = Val(txtPenalty.Text)
348                                     !Passbook = Val(txtPassbook.Text)
350                                     !TotalAmortization = Val( _
                                                txtTotalAmortization.Text)
352                                     !CollectorCode = lblCollectorCode.Caption
354                                     !CollectorFname = lblCollectorFname.Caption
356                                     !TotalCharges = Val(txtTotalCharges.Text)
358                                     !LoanTotal = Val(txtLoanTotal.Text)
360                                     !User = lblUser.Caption
362                                     !LoanStatus = "Good"
364                                     !TotalPayment = Val(txtCollection.Text)
366                                     !NotPosted = Val(txtNotPosted.Text)
368                                     !OverToCus = Val(txtOverToCus.Text)
                                        
                                        Select Case OptType(0)

                                            Case True ' Regular Loan
                                                LoanTypeVar = 1 '"Regular"

                                            Case False ' Emergency
                                                 
                                                LoanTypeVar = 2 '"Emergency"
                                        End Select

                                        !LOANTYPE = LoanTypeVar
                                        !LoanPeriod = Val(txtTerms)
                                        !InterestRate = Val(txtInterestRate)
                                        
370                                     !PaymentPerDay = Val(lblPerDay.Caption)
372                                     .Update
                                    End With
                    
                                End If
                        
                                'Updates the New Balance of the Customer
374                             If rsCustomer.State = 1 Then rsCustomer.Close
376                             rsCustomer.Open _
                                        "Select * from tblCustomer where Code = " & _
                                        cCode & ""

                                'in some instances it does not update. hmn why?
                                Select Case OptType(0)

                                    Case True ' Regular Loan
378                                     rsCustomer!Balance = (rsLoan!TotalAmortization) _
                                                'disk or network error. wow.

                                    Case False ' Emergency
                                     rsCustomer!EMERGENCYBalance = ( _
                                                rsLoan!TotalAmortization) 'disk or network error. wow.
                                End Select
                                
380                             rsCustomer.Update
                    
382                             If rsTrail.State = 1 Then rsTrail.Close
384                             rsTrail.Open "Select * from tblTrail "

                                'Records when adding a Loan
386                             With rsTrail
388                                 .AddNew
390                                 !UserName = lblUser.Caption
392                                 !userlevel = lblUserlevel.Caption
394                                 !Activity = "Add New Loan"
396                                 !Time = lblTime.Caption
398                                 !Date = lblDate.Caption
400                                 .Update
                                End With
                            
402                             MsgBox "New Loan successfully created!", vbInformation, _
                                        "Webplus Lending Corporation"
404                             lblBalance.Caption = "0"

406                             btnClose_Click
408                             frm_Loan.lblUser.Caption = MDIForm1.lblusername.Caption
410                             frm_Loan.lblUserlevel.Caption = _
                                        MDIForm1.lblUserlevel.Caption
412                             DataGrid1.Refresh
414                             Me.Show
                            
416                             btnClose.Caption = "&Close"
                       
418                             txtSearch.Enabled = True
                            End If

                            'Print lblCuCode.Caption
                       
420                         Print cCode

                            'update latest loan totalammortization to costumer balance
422                         If rsLoan1.State = 1 Then rsLoan1.Close
424                         rsLoan1.Open " Select * from tblLoan where Code = " & cCode _
                                    & " and Status = 'Good' Order By LoanID Desc"
                        
426                         If rsLoan1.RecordCount <> 0 Then
428                             rsLoan1.MoveFirst 'move to loan with highest id
                        
430                             If rsCustomer2.State = 1 Then rsCustomer2.Close
432                             rsCustomer2.Open _
                                        "Select * from tblCustomer where Code = " & _
                                        cCode & ""
                        
434                             rsCustomer2!Balance = rsLoan1!TotalAmortization
                            
436                             rsCustomer2.Update
                        
438                             rsCustomer2.Close
440                             rsLoan1.Close
                            End If
                        
442                         cCode = 0

                            'nag yes jud siya sa payment diri.
                        End If
                    
                    End If
                End If
            End If
        End If

444     Call txtTerms_Change

        '<EhFooter>

        TxtLog "Exited btnAdd_Click"

        Exit Sub

btnAdd_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.btnAdd_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       btnClose_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub btnClose_Click()

        '<EhHeader>
        On Error GoTo btnClose_Click_Err

        TxtLog "Entered btnClose_Click"

        '</EhHeader>

100     If btnClose.Caption = "&Close" Then
102         Unload Me
104         MDIForm1.Picture1.Visible = True
        Else
106         txtSearch.Enabled = True
108         btnClose.Caption = "&Close"
110         btnAdd.Caption = "&Add"
112         txtCode.Text = ""
114         txtCollector.Text = ""
116         txtServicefee.Text = "0"
118         txtLifeInsurance.Text = ""
120         txtMature.Text = ""
122         txtFirstName.Text = ""
124         txtFCollector.Text = ""
126         txtPrincipal.Text = "0"
128         txtTotal.Text = ""
130         txtCustomer.Text = ""
132         txtFireInsurance.Text = ""
134         txtCollection.Text = "0"
136         txtServicefee.Text = "0"
138         txtDelivery.Text = "0"
140         txtPenalty.Text = ""
142         txtPassbook.Text = ""
144         txtCharge.Text = "0"
146         dtDateRelease.Enabled = False
148         txtCode.Enabled = False
150         txtPrincipal.Enabled = False
152         Call sort
154         Call Timer1_Timer
            
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       DataGrid1_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
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
        ErrReport Err.Description, "LendingClient.frm_Loan.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       dtDaterelease_Change
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub dtDaterelease_Change()

        '<EhHeader>
        On Error GoTo dtDaterelease_Change_Err

        TxtLog "Entered dtDaterelease_Change"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     txtTerms_Change

104     If rsLoan.State = 1 Then rsLoan.Close
106     rsLoan.Open "Select * from tblLoan where DateRelease =  #" & _
                dtDateRelease.Value & "# and Code = " & Val(lblCuCode.Caption) & _
                " and Status = '" & "Good" & "'"

        ' dada = rsLoan!DateRelease
108     If rsLoan.RecordCount <> 0 Then
110         MsgBox "This Customer is already loaned on " & rsLoan!DateRelease
112         txtCode.SetFocus
114         Call Timer1_Timer
        Else
        End If

        'txtMature.Text = DateAdd("m", 2, dtDaterelease.Value)
                  
116     If rsCollData.State = 1 Then rsCollData.Close
118     rsCollData.Open "Select * from tblColl_Data where Code = " & _
                Val(lblCollectorCode.Caption) & " and  DateEmployed <= #" & _
                dtDateRelease.Value & "# order by DateEmployed DESC"
                
120     If rsCollData.RecordCount <> 0 Then
122         rsCollData.MoveFirst
124         lblCollectorFname.Caption = rsCollData!firstname
126         txtFCollector.Text = rsCollData!firstname
128         txtCollector.Text = rsCollData!lastname
        Else
130         lblCollectorFname.Caption = ""
132         txtFCollector.Text = ""
134         txtCollector.Text = ""
        End If

        '<EhFooter>

        TxtLog "Exited dtDaterelease_Change"

        Exit Sub

dtDaterelease_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.dtDaterelease_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       dtDaterelease_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub dtDaterelease_Click()

        '<EhHeader>
        On Error GoTo dtDaterelease_Click_Err

        TxtLog "Entered dtDaterelease_Click"

        '</EhHeader>

        'ok theb
        'Call Timer1_Timer
        '  dtDaterelease.Value = DateAdd("m", 2, Date)

100     txtMature.Text = DateAdd("m", 2, dtDateRelease.Value)

        '<EhFooter>

        TxtLog "Exited dtDaterelease_Click"

        Exit Sub

dtDaterelease_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.dtDaterelease_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       dtDaterelease_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub dtDaterelease_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo dtDaterelease_KeyPress_Err

        TxtLog "Entered dtDaterelease_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         txtFireInsurance.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited dtDaterelease_KeyPress"

        Exit Sub

dtDaterelease_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.dtDaterelease_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Form_Load
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

        '
100     Call txtTerms_Change
102     Call connect
104     Call Servicefee
        '     Call delivery
106     Call Customer
108     Call payment
110     Call Loan
112     Call Trail
114     Call Collector
116     Call Loan1
118     Call Customer2
120     Call CollCode
122     Call CollData
124     Me.Show
126     lblLifeInsurance.Visible = False

128     Me.SetFocus

        'Me.WindowState = vbNormal
130     If rsLoan.State = 1 Then rsLoan.Close
        'Sort records to descending
132     rsLoan.Open "Select * from tblLoan Order By LoanID desc"
134     Set DataGrid1.DataSource = rsLoan
136     DataGrid1.Columns(5).Caption = "Encoded Date"
138     DataGrid1.Width = Me.Width
140     DataGrid1.Columns(1).Width = 1100
142     DataGrid1.Columns(2).Width = 1100
144     DataGrid1.Columns(4).Width = 1500
146     DataGrid1.Columns(5).Width = 1150
148     DataGrid1.Columns(6).Width = 750
150     DataGrid1.Columns(7).Width = 1000
152     DataGrid1.Columns(8).Width = 1300
154     DataGrid1.Columns(9).Width = 1100
156     DataGrid1.Columns(10).Width = 1100
158     DataGrid1.Columns(11).Width = 1100
160     DataGrid1.Columns(22).Width = 880
        '158     With Me
        '160         .Top = (MDIForm1.Height - .Height)
        '162         .Left = (MDIForm1.Width - .Width)
        '        End With
162     Me.WindowState = vbMaximized
        ' txtDelivery.Text = rsDelivery!delivery
164     txtDelivery.Text = 0
    
166     txtP.Text = Val(txtPrincipal.Text)
        lblPerDay.Caption = "0"
168     btnAdd.SetFocus

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub OptMonth_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo OptMonth_Click_Err

        TxtLog "Entered OptMonth_Click"

        '</EhHeader>

100     Select Case Index

            Case 0
102             txtTerms = 30

104         Case 1
106             txtTerms = 45

108         Case 2
110             txtTerms = 60

112         Case 3
114             txtTerms = 75

116         Case 4
118             txtTerms = 90
        
        End Select

        '<EhFooter>

        TxtLog "Exited OptMonth_Click"

        Exit Sub

OptMonth_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.OptMonth_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub OptType_Click(Index As Integer)

        '<EhHeader>
        On Error GoTo OptType_Click_Err

        TxtLog "Entered OptType_Click"

        '</EhHeader>

100     Select Case Index
        
            Case 0 'Regular Loan
            
102             Period.Visible = 1
104             txtInterestRate = Val(6)
106             txtTerms.Visible = 1
108             txtMature.Visible = 1
110             Frame2.Visible = True
112             Period.Visible = True
114             txtTerms = 30
116             OptMonth(0) = True

118         Case 1 ' Emergency
120             Period.Visible = False
122             txtInterestRate = Val(1)
124             txtTerms.Visible = True
126             txtMature.Visible = False
                
                'Initialize fields for emergency loan type
128             Frame2.Visible = False
130             Period.Visible = False
132             txtTerms = 1

134             If OptType(1).Value = True Then txtTotalCharges.Text = 0
136             txtLoanTotal.Text = Val(txtPrincipal)
        End Select

        '<EhFooter>

        TxtLog "Exited OptType_Click"

        Exit Sub

OptType_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.OptType_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Timer1_Timer
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>
   
    ' dim NextMonth as Date  dateadd(Dateinterval.month, 2, #8/31/2014#)

    lblDate.Caption = FormatDateTime(Date, vbShortDate)
    lblTime.Caption = Time
    ' mature = DateAdd("m", 2, Date)
    ' txtMature.Text = mature
    ' dtDaterelease.Value = Date
    '  dtDaterelease.Value = DateAdd("m", 2, Date)
    'dtDaterelease.Value = DateAdd("m", 1, Date)
    'txtMature.Text = dtDaterelease.Value
   
    '
    'dtmaturity.Value = mature
    dtDateRelease.Value = Date
    ' DateAdd (dateinterval.month , 1 , dtmature.value)
    Timer1.Enabled = False
 
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtCharge_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtCharge_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtCharge_KeyPress_Err

        TxtLog "Entered txtCharge_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf KeyAscii = 13 Then
        
        Else
110         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtCharge_KeyPress"

        Exit Sub

txtCharge_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtCharge_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtCharge_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtCharge_LostFocus()

        '<EhHeader>
        On Error GoTo txtCharge_LostFocus_Err

        TxtLog "Entered txtCharge_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     Call SumCharges

        '<EhFooter>

        TxtLog "Exited txtCharge_LostFocus"

        Exit Sub

txtCharge_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtCharge_LostFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtCode_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
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
            
110         Call lenth
            
            'add also this code to code textbox to check immedietly if the customer record is reversed.
            'Recheck Computation
            'Mao ne ang mo check if ang customer gereverse na ba siya..
            'Kung gereverse na gane dle na siya makaloan ug balik
            'Pero ug wala pa ge reverse makaloan pa siya
128         If rsCustomer.State = 1 Then rsCustomer.Close
130         rsCustomer.Open "Select * from tblCustomer where Code = " & Val( _
                    txtCode.Text) & " and status = 'Reversed'"

132         If rsCustomer.RecordCount > 0 Then
134             MsgBox "Customer is already reversed. He/she cannot loan."
                txtCode.SetFocus

                Exit Sub
                
            End If
            
            'Searching for Existing Customer record based on Customer Code
112         If rsCustomer.State = 1 Then rsCustomer.Close
114         rsCustomer.Open "Select * from tblCustomer where Code = " & Val( _
                    txtCode.Text) & ""

            'Validate if customer exist then take action
116         If rsCustomer.RecordCount < 1 Then
                If txtCode.Text = "" Then Exit Sub
118             MsgBox "Customer record not found!"

                Exit Sub

            End If
            
            'Search if there is existing loan based on daterelease
120         If rsLoan.State = 1 Then rsLoan.Close
122         rsLoan.Open "Select * from tblLoan where DateRelease = #" & _
                    dtDateRelease.Value & "# and Code = " & txtCode.Text & _
                    " and Status = '" & "Good" & "'"
                
124         If rsCustomer.RecordCount = 0 Then
126             MsgBox "No Record Found !", vbInformation
                txtCode.Text = ""
                'If there is existing loan today, can't add more on the same date
            ElseIf rsLoan.RecordCount <> 0 Then
                MsgBox "This Customer is already loaned Today  Can't add more loan "
                txtCode.Text = ""
            Else

                'If the Customer record exist
136             If OptType(0) = True Then
                Else
138                 txtPrincipal.SetFocus

                    Exit Sub

                End If

140             With rsCustomer
                   
                    'txtCollector.Text = !Collector
142                 txtCustomer.Text = !lastname
144                 txtFirstName.Text = !firstname
146                 txtLifeInsurance.Text = !Balance
148                 lblLifeInsurance.Caption = !Balance
150                 lblBalance.Caption = !Balance
152                 lblCollectorCode.Caption = !CollectorCode
                    'lblCollectorFname.Caption = !CollectorFirstname
154                 lblCuCode.Caption = !code 'Alera
                    'txtFCollector.Text = !CollectorFirstname

                End With
                   
156             If rsCollData.State = 1 Then rsCollData.Close
158             rsCollData.Open "Select * from tblColl_Data where Code = " & _
                        lblCollectorCode.Caption & " and  DateEmployed <= #" & _
                        dtDateRelease.Value & "# order by DateEmployed DESC"
                
160             If rsCollData.RecordCount <> 0 Then
162                 rsCollData.MoveFirst
164                 lblCollectorFname.Caption = rsCollData!firstname
166                 txtFCollector.Text = rsCollData!firstname
168                 txtCollector.Text = rsCollData!lastname
                Else
170                 lblCollectorFname.Caption = ""
172                 txtFCollector.Text = ""
174                 txtCollector.Text = ""
                End If
                
                'Pointing to Principal Field
176             txtPrincipal.SetFocus
                
            End If
            
        Else
178         KeyAscii = 0
        End If

        'End If

        '<EhFooter>

        TxtLog "Exited txtCode_KeyPress"

        Exit Sub

txtCode_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtCode_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtCode_LostFocus()
    Call txtCode_KeyPress(13)

End Sub

Private Sub txtCollection_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtCollection_KeyPress_Err

        TxtLog "Entered txtCollection_KeyPress"

        '</EhHeader>

        Dim strvalid As String
        
100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
        
108     ElseIf txtCollection.Text = "" Then
110         MsgBox "Please put the exact amount.", vbInformation
112         txtCollection.SetFocus
114     ElseIf KeyAscii = 13 Then
116         txtServicefee.SetFocus
        Else
118         KeyAscii = 0
    
        End If

        '<EhFooter>

        TxtLog "Exited txtCollection_KeyPress"

        Exit Sub

txtCollection_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtCollection_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtCollection_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtCollection_LostFocus()

        '<EhHeader>
        On Error GoTo txtCollection_LostFocus_Err

        TxtLog "Entered txtCollection_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     If txtCollection.Text = "" Then
        Else
104         Call SumCharges
        End If

        '<EhFooter>

        TxtLog "Exited txtCollection_LostFocus"

        Exit Sub

txtCollection_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtCollection_LostFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtCollection_MouseDown
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       Button (Integer)
'                    Shift (Integer)
'                    x (Single)
'                    Y (Single)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtCollection_MouseDown(Button As Integer, _
                                    Shift As Integer, _
                                    x As Single, _
                                    Y As Single)

        '<EhHeader>
        On Error GoTo txtCollection_MouseDown_Err

        TxtLog "Entered txtCollection_MouseDown"

        '</EhHeader>

100     If Button = vbRightButton Then
102         txtCollection.Locked = True
        Else
104         txtCollection.Locked = False
        End If

        '<EhFooter>

        TxtLog "Exited txtCollection_MouseDown"

        Exit Sub

txtCollection_MouseDown_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtCollection_MouseDown", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtDelivery_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtDelivery_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtDelivery_KeyPress_Err

        TxtLog "Entered txtDelivery_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         txtPenalty.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtDelivery_KeyPress"

        Exit Sub

txtDelivery_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtDelivery_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtDelivery_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtDelivery_LostFocus()

        '<EhHeader>
        On Error GoTo txtDelivery_LostFocus_Err

        TxtLog "Entered txtDelivery_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     Call SumCharges

        '<EhFooter>

        TxtLog "Exited txtDelivery_LostFocus"

        Exit Sub

txtDelivery_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtDelivery_LostFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtDelivery_MouseDown
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       Button (Integer)
'                    Shift (Integer)
'                    x (Single)
'                    Y (Single)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtDelivery_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  Y As Single)

        '<EhHeader>
        On Error GoTo txtDelivery_MouseDown_Err

        TxtLog "Entered txtDelivery_MouseDown"

        '</EhHeader>

100     If Button = vbRightButton Then
102         txtDelivery.Locked = True
        Else
104         txtDelivery.Locked = False
        End If

        '<EhFooter>

        TxtLog "Exited txtDelivery_MouseDown"

        Exit Sub

txtDelivery_MouseDown_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtDelivery_MouseDown", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtFireInsurance_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtFireInsurance_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtFireInsurance_KeyPress_Err

        TxtLog "Entered txtFireInsurance_KeyPress"

        '</EhHeader>

        Dim strvalid As String
        
100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf txtFireInsurance.Text = "" Then
110         MsgBox "Please put the exact amount.", vbInformation
112         txtFireInsurance.SetFocus
114     ElseIf KeyAscii = 13 Then
116         Call SumCharges
            'txtCollection.Enabled = True
            '120         txtCollection.SetFocus
        Else
118         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtFireInsurance_KeyPress"

        Exit Sub

txtFireInsurance_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtFireInsurance_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtFireInsurance_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtFireInsurance_LostFocus()

        '<EhHeader>
        On Error GoTo txtFireInsurance_LostFocus_Err

        TxtLog "Entered txtFireInsurance_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub
        
102     If txtFireInsurance.Text = "" Then
        Else

104         If btnAdd.Caption = "&Add" Then

                Exit Sub

            End If

106         Call SumCharges
        End If

        '<EhFooter>

        TxtLog "Exited txtFireInsurance_LostFocus"

        Exit Sub

txtFireInsurance_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtFireInsurance_LostFocus", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtFireInsurance_MouseDown
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       Button (Integer)
'                    Shift (Integer)
'                    x (Single)
'                    Y (Single)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtFireInsurance_MouseDown(Button As Integer, _
                                       Shift As Integer, _
                                       x As Single, _
                                       Y As Single)

        '<EhHeader>
        On Error GoTo txtFireInsurance_MouseDown_Err

        TxtLog "Entered txtFireInsurance_MouseDown"

        '</EhHeader>

100     If Button = vbRightButton Then
102         txtFireInsurance.Locked = True
        Else
104         txtFireInsurance.Locked = False
        End If

        '<EhFooter>

        TxtLog "Exited txtFireInsurance_MouseDown"

        Exit Sub

txtFireInsurance_MouseDown_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtFireInsurance_MouseDown", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtInterestRate_Change()

        '<EhHeader>
        On Error GoTo txtInterestRate_Change_Err

        TxtLog "Entered txtInterestRate_Change"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     Call SumCharges
104     txtPrincipal_Change

        '<EhFooter>

        TxtLog "Exited txtInterestRate_Change"

        Exit Sub

txtInterestRate_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtInterestRate_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtInterestRate_GotFocus()

        '<EhHeader>
        On Error GoTo txtInterestRate_GotFocus_Err

        TxtLog "Entered txtInterestRate_GotFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     Call SumCharges
104     txtPrincipal_Change

        '<EhFooter>

        TxtLog "Exited txtInterestRate_GotFocus"

        Exit Sub

txtInterestRate_GotFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtInterestRate_GotFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtInterestRate_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtInterestRate_KeyPress_Err

        TxtLog "Entered txtInterestRate_KeyPress"

        '</EhHeader>

100     Call SumCharges
102     txtPrincipal_Change

        '<EhFooter>

        TxtLog "Exited txtInterestRate_KeyPress"

        Exit Sub

txtInterestRate_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtInterestRate_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtInterestRate_LostFocus()

        '<EhHeader>
        On Error GoTo txtInterestRate_LostFocus_Err

        TxtLog "Entered txtInterestRate_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     Call SumCharges

        '<EhFooter>

        TxtLog "Exited txtInterestRate_LostFocus"

        Exit Sub

txtInterestRate_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtInterestRate_LostFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtLifeInsurance_Change()

        '<EhHeader>
        On Error GoTo txtLifeInsurance_Change_Err

        TxtLog "Entered txtLifeInsurance_Change"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then

            Exit Sub

        End If

102     If Val(txtLifeInsurance.Text) <> 0 Then
104         OptType(1).Visible = True
        Else
106         OptType(1).Visible = False
        End If
    
        '<EhFooter>

        TxtLog "Exited txtLifeInsurance_Change"

        Exit Sub

txtLifeInsurance_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtLifeInsurance_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtLifeInsurance_KeyPress
' Description:       current regular loan balance
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtLifeInsurance_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtLifeInsurance_KeyPress_Err

        TxtLog "Entered txtLifeInsurance_KeyPress"

        '</EhHeader>

        Dim strvalid As String
        
100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf txtLifeInsurance.Text = "" Then
110         MsgBox "Please put the exact amount.", vbInformation
112         txtLifeInsurance.SetFocus
114     ElseIf KeyAscii = 13 Then
116         txtPenalty.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtLifeInsurance_KeyPress"

        Exit Sub

txtLifeInsurance_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtLifeInsurance_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtLifeInsurance_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtLifeInsurance_LostFocus()

        '<EhHeader>
        On Error GoTo txtLifeInsurance_LostFocus_Err

        TxtLog "Entered txtLifeInsurance_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     If txtLifeInsurance.Text = "" Then
        Else
104         Call SumCharges
        End If

        '<EhFooter>

        TxtLog "Exited txtLifeInsurance_LostFocus"

        Exit Sub

txtLifeInsurance_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtLifeInsurance_LostFocus", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtNotPosted_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtNotPosted_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtNotPosted_KeyPress_Err

        TxtLog "Entered txtNotPosted_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf KeyAscii = 13 Then
        
        Else
110         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtNotPosted_KeyPress"

        Exit Sub

txtNotPosted_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtNotPosted_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtNotPosted_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtNotPosted_LostFocus()

        '<EhHeader>
        On Error GoTo txtNotPosted_LostFocus_Err

        TxtLog "Entered txtNotPosted_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     Call SumCharges

        '<EhFooter>

        TxtLog "Exited txtNotPosted_LostFocus"

        Exit Sub

txtNotPosted_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtNotPosted_LostFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtOverToCus_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtOverToCus_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtOverToCus_KeyPress_Err

        TxtLog "Entered txtOverToCus_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf KeyAscii = 13 Then
        
        Else
110         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtOverToCus_KeyPress"

        Exit Sub

txtOverToCus_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtOverToCus_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtOverToCus_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtOverToCus_LostFocus()

        '<EhHeader>
        On Error GoTo txtOverToCus_LostFocus_Err

        TxtLog "Entered txtOverToCus_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     Call SumCharges

        '<EhFooter>

        TxtLog "Exited txtOverToCus_LostFocus"

        Exit Sub

txtOverToCus_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtOverToCus_LostFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtPassbook_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPassbook_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtPassbook_KeyPress_Err

        TxtLog "Entered txtPassbook_KeyPress"

        '</EhHeader>

        Dim strvalid As String
        
100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf KeyAscii = 13 Then
            'Call SumCharges
110         Call btnAdd.SetFocus
        Else
112         KeyAscii = 13
        End If

        '<EhFooter>

        TxtLog "Exited txtPassbook_KeyPress"

        Exit Sub

txtPassbook_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtPassbook_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtPassbook_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPassbook_LostFocus()

        '<EhHeader>
        On Error GoTo txtPassbook_LostFocus_Err

        TxtLog "Entered txtPassbook_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     Call SumCharges

        '<EhFooter>

        TxtLog "Exited txtPassbook_LostFocus"

        Exit Sub

txtPassbook_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtPassbook_LostFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtPassbook_MouseDown
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       Button (Integer)
'                    Shift (Integer)
'                    x (Single)
'                    Y (Single)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPassbook_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  Y As Single)

        '<EhHeader>
        On Error GoTo txtPassbook_MouseDown_Err

        TxtLog "Entered txtPassbook_MouseDown"

        '</EhHeader>

100     If Button = vbRightButton Then
102         txtPassbook.Locked = True
        Else
104         txtPassbook.Locked = False
        End If

        '<EhFooter>

        TxtLog "Exited txtPassbook_MouseDown"

        Exit Sub

txtPassbook_MouseDown_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtPassbook_MouseDown", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtPenalty_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPenalty_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtPenalty_KeyPress_Err

        TxtLog "Entered txtPenalty_KeyPress"

        '</EhHeader>

        Dim strvalid As String
        
100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf KeyAscii = 13 Then
            '  Sum = Val(txtFireInsurance.Text) + Val(txtCollection.Text) + Val(txtServicefee.Text) + Val(txtLifeInsurance.Text) + Val(txtDelivery.Text) + Val(txtPenalty.Text) + Val(txtPassbook.Text)
            '
            ' Summ = Val(txtPrincipal.Text) - Val(txtTotalCharges.Text)
            '  txtLoanTotal.Text = Val(Summ)
110         txtPassbook.SetFocus
        Else
112         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtPenalty_KeyPress"

        Exit Sub

txtPenalty_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtPenalty_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtPenalty_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPenalty_LostFocus()

        '<EhHeader>
        On Error GoTo txtPenalty_LostFocus_Err

        TxtLog "Entered txtPenalty_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     Call SumCharges
104     txtPassbook.SetFocus

        '<EhFooter>

        TxtLog "Exited txtPenalty_LostFocus"

        Exit Sub

txtPenalty_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtPenalty_LostFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtPenalty_MouseDown
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       Button (Integer)
'                    Shift (Integer)
'                    x (Single)
'                    Y (Single)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPenalty_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 Y As Single)

        '<EhHeader>
        On Error GoTo txtPenalty_MouseDown_Err

        TxtLog "Entered txtPenalty_MouseDown"

        '</EhHeader>

100     If Button = vbRightButton Then
102         txtPenalty.Locked = True
        Else
104         txtPenalty.Locked = False
        End If

        '<EhFooter>

        TxtLog "Exited txtPenalty_MouseDown"

        Exit Sub

txtPenalty_MouseDown_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtPenalty_MouseDown", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtPrincipal_Change
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPrincipal_Change()

        '<EhHeader>
        On Error GoTo txtPrincipal_Change_Err

        TxtLog "Entered txtPrincipal_Change"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub
102     If Val(txtPrincipal.Text) = 0 Then Exit Sub

        Dim Interest As Double

        Dim Balance  As Double
       
        'Computing for the total balanced based on how much loan
        'CSBmk <computations>
104     Interest = Val(txtPrincipal.Text) * (Val(txtInterestRate) / 100)
106     Balance = Val(txtPrincipal.Text) + Val(Interest)
108     txtTotal.Text = Val(Balance)
110     txtTotalAmortization = Val(txtTotal.Text)
112     txtP = txtPrincipal

114     Select Case txtInterestRate
        
            Case 6
116             txtFireInsurance.Text = Val(txtPrincipal) * 0.02

118         Case 9
120             txtFireInsurance.Text = Val(txtPrincipal) * 0.03

122         Case 12
124             txtFireInsurance.Text = Val(txtPrincipal) * 0.04

126         Case 15
128             txtFireInsurance.Text = Val(txtPrincipal) * 0.05

130         Case 18
132             txtFireInsurance.Text = Val(txtPrincipal) * 0.06
        
        End Select
        
        Call SumCharges

        '<EhFooter>

        TxtLog "Exited txtPrincipal_Change"

        Exit Sub

txtPrincipal_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtPrincipal_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtPrincipal_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPrincipal_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtPrincipal_KeyPress_Err

        TxtLog "Entered txtPrincipal_KeyPress"

        '</EhHeader>

        Dim strvalid As String

        'Dim RRCode   As String /// delete later

100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf KeyAscii = 13 Then
            'RRCode = rsCustomer!code

            '116         If rsCustomer.State = 1 Then rsCustomer.Close
            '118         rsCustomer.Open "Select Balance from tblCustomer where Code = '" & RRCode & _
            '                    "'"
            '
            '120         If rsCustomer.RecordCount = 0 Then
            '122             MsgBox "No Record Found", vbInformation
            '            Else
            '
            '124             If rsCustomer.State = 1 Then rsCustomer.Close
            '126             rsCustomer.Open "Select *  from tblCustomer where Code = '" & RRCode & "'"
            '128             txtTotalAmortization.Text = txtTotal.Text
            '            End If
            '
110         txtP.Text = Val(txtPrincipal.Text)
112         dtDateRelease.SetFocus
        Else
114         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtPrincipal_KeyPress"

        Exit Sub

txtPrincipal_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtPrincipal_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtPrincipal_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPrincipal_LostFocus()

        '<EhHeader>
        On Error GoTo txtPrincipal_LostFocus_Err

        TxtLog "Entered txtPrincipal_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub
 
102     txtP.Text = txtPrincipal.Text
    
104     Call SumCharges

        '<EhFooter>

        TxtLog "Exited txtPrincipal_LostFocus"

        Exit Sub

txtPrincipal_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtPrincipal_LostFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtPrincipal_MouseDown
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       Button (Integer)
'                    Shift (Integer)
'                    x (Single)
'                    Y (Single)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtPrincipal_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, _
                                   Y As Single)

        '<EhHeader>
        On Error GoTo txtPrincipal_MouseDown_Err

        TxtLog "Entered txtPrincipal_MouseDown"

        '</EhHeader>

100     If Button = vbRightButton Then
102         txtPrincipal.Locked = True
        Else
104         txtPrincipal.Locked = False
        End If

        '<EhFooter>

        TxtLog "Exited txtPrincipal_MouseDown"

        Exit Sub

txtPrincipal_MouseDown_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtPrincipal_MouseDown", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtsearch_Change
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsLoan.State = 1 Then rsLoan.Close
102     rsLoan.Open "Select * from tblLoan where Customer like '" & txtSearch.Text & _
                "%' or Code like '" & txtSearch.Text & "%' or LoanID like '" & _
                txtSearch.Text & "%' or FirstName like '" & txtSearch.Text & _
                "%' Order by LoanID Desc"
        
104     Set DataGrid1.DataSource = rsLoan

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtServicefee_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtServicefee_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtServicefee_KeyPress_Err

        TxtLog "Entered txtServicefee_KeyPress"

        '</EhHeader>

        Dim strvalid As String
        
100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf txtServicefee.Text = "" Then
110         MsgBox "Please put the exact amount.", vbInformation
112         txtServicefee.SetFocus
114     ElseIf KeyAscii = 13 Then
116         txtPenalty.SetFocus
        Else
118         KeyAscii = 0
            
        End If

        '<EhFooter>

        TxtLog "Exited txtServicefee_KeyPress"

        Exit Sub

txtServicefee_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtServicefee_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtServicefee_LostFocus
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtServicefee_LostFocus()

        '<EhHeader>
        On Error GoTo txtServicefee_LostFocus_Err

        TxtLog "Entered txtServicefee_LostFocus"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

102     If txtServicefee.Text = "" Then
        Else
104         Call SumCharges
        End If

        '<EhFooter>

        TxtLog "Exited txtServicefee_LostFocus"

        Exit Sub

txtServicefee_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtServicefee_LostFocus", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtTerms_Change()

        '<EhHeader>
        On Error GoTo txtTerms_Change_Err

        TxtLog "Entered txtTerms_Change"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then Exit Sub

        'SumCharges
102     If OptType(0).Value = True Then
            
        Else

            Exit Sub

        End If

104     txtPrincipal_Change
    'todo add exit sub condition if no principal yet
106     Select Case txtTerms
        
            Case 30 '1 month
108             txtInterestRate = 6
110             txtMature.Text = DateAdd("d", 30, dtDateRelease.Value)

112         Case 45 '1.5 months
114             txtInterestRate = 9
116             txtMature.Text = DateAdd("d", 45, dtDateRelease.Value)

118         Case 60 '2 months
120             txtInterestRate = 12
122             txtMature.Text = DateAdd("d", 60, dtDateRelease.Value)

124         Case 75 '2.5 months
126             txtInterestRate = 15
128             txtMature.Text = DateAdd("d", 75, dtDateRelease.Value)

130         Case 90 ' 3 months
132             txtInterestRate = 18
134             txtMature.Text = DateAdd("d", 90, dtDateRelease.Value)

136         Case Else
                'MsgBox ("Invalid value please type again")
138             txtInterestRate = "invalid"
    
        End Select

        '<EhFooter>

        TxtLog "Exited txtTerms_Change"

        Exit Sub

txtTerms_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtTerms_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       txtTotal_KeyPress
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:00 AM
'
' Parameters :       KeyAscii (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub txtTotal_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtTotal_KeyPress_Err

        TxtLog "Entered txtTotal_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
  
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8

        Else
108         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtTotal_KeyPress"

        Exit Sub

txtTotal_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Loan.txtTotal_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

