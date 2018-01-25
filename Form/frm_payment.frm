VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_payment 
   Caption         =   "frmPayment"
   ClientHeight    =   10725
   ClientLeft      =   4305
   ClientTop       =   -435
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10725
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Payment Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10455
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   20295
      Begin VB.TextBox txtFCollector 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   54
         Top             =   1200
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   7680
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Format          =   265027585
         CurrentDate     =   41894
      End
      Begin VB.TextBox txtSearch 
         Height          =   405
         Left            =   1920
         TabIndex        =   7
         Top             =   5280
         Width           =   3735
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   5760
         Width           =   20055
         _ExtentX        =   35375
         _ExtentY        =   3625
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
      Begin VB.TextBox txtLoanID 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   4560
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtFirstname 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   4560
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtOR 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   2520
         TabIndex        =   6
         Top             =   7920
         Width           =   1815
      End
      Begin VB.CommandButton btnAddpayment 
         Caption         =   "&Add Payment"
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   7920
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   5880
         Top             =   3240
      End
      Begin VB.TextBox txtTodayDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   9720
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   3255
      End
      Begin VB.Frame Frame2 
         Caption         =   "Balance Info"
         Height          =   2175
         Left            =   7200
         TabIndex        =   30
         Top             =   600
         Width           =   6135
         Begin VB.TextBox txtTotaltBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2640
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1320
            Width           =   3135
         End
         Begin VB.TextBox txtPaymentmade 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2640
            TabIndex        =   14
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox txtTotaltAmortization 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   405
            Left            =   2640
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   240
            Width           =   3135
         End
         Begin VB.Line Line3 
            X1              =   1920
            X2              =   5760
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line2 
            X1              =   1920
            X2              =   5760
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label Label12 
            Caption         =   "Payments made  :  "
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
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label11 
            Caption         =   "Amortization        :"
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
            TabIndex        =   34
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total balance         :"
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
            Left            =   120
            TabIndex        =   33
            Top             =   1320
            Width           =   2385
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   5880
            Y1              =   1200
            Y2              =   1200
         End
      End
      Begin VB.TextBox txtMaturity 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4440
         Width           =   3855
      End
      Begin VB.TextBox txtDaterelease 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   1920
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3960
         Width           =   3855
      End
      Begin VB.TextBox txtAmortization 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox txtPrincipal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtCustomer 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox txtCollector 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Frame Frame3 
         Caption         =   "Section"
         Height          =   2175
         Left            =   7200
         TabIndex        =   31
         Top             =   2880
         Width           =   8535
         Begin VB.TextBox txtTotalBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2880
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   240
            Width           =   2775
         End
         Begin VB.TextBox txtAmountPaid 
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
            Height          =   675
            Left            =   2880
            MaxLength       =   5
            TabIndex        =   2
            Top             =   660
            Width           =   2775
         End
         Begin VB.TextBox txtNewBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   405
            Left            =   2880
            TabIndex        =   16
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1560
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker dtDate 
            Height          =   375
            Left            =   6120
            TabIndex        =   3
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   265027585
            CurrentDate     =   41869
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Payment:"
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
            Left            =   6120
            TabIndex        =   47
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Outstanding  Balance  :"
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
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   2475
         End
         Begin VB.Line Line5 
            X1              =   5760
            X2              =   2880
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line4 
            X1              =   5640
            X2              =   2880
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Less: Amount Paid"
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
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   1995
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Total Balance    :"
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
            Left            =   120
            TabIndex        =   36
            Top             =   1560
            Width           =   1755
         End
      End
      Begin VB.Line Line6 
         BorderStyle     =   4  'Dash-Dot
         X1              =   120
         X2              =   17400
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label_NewOr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   14160
         TabIndex        =   61
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblPaymentsMade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   6240
         TabIndex        =   60
         Top             =   4320
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblOver 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   13680
         TabIndex        =   59
         Top             =   2280
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblCuCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5520
         TabIndex        =   58
         Top             =   2160
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lbltt 
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
         Left            =   9840
         TabIndex        =   57
         Top             =   3720
         Width           =   75
      End
      Begin VB.Label lblLabel1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search using customer name, ORnumber, Customer Code or Encoded Date"
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
         Left            =   5880
         TabIndex        =   56
         Top             =   5400
         Width           =   6720
      End
      Begin VB.Label lblAuOR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6120
         TabIndex        =   55
         Top             =   4320
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblCollectorFname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ppprte"
         Height          =   195
         Left            =   4680
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblCollectorCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CollectorCode"
         Height          =   195
         Left            =   13800
         TabIndex        =   52
         Top             =   1200
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   4440
         TabIndex        =   51
         Top             =   2040
         Width           =   90
      End
      Begin VB.Label lblCodeMust 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note:       Code must only be typed by numbers"
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
         Left            =   1080
         TabIndex        =   50
         Top             =   1680
         Width           =   4110
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kkkkkkkkk"
         Height          =   195
         Left            =   5880
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblSearchHere 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Here: "
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
         TabIndex        =   46
         Top             =   5280
         Width           =   1470
      End
      Begin VB.Label erer 
         AutoSize        =   -1  'True
         Caption         =   "Loan ID            :"
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
         Left            =   2760
         TabIndex        =   45
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "OR Number     :"
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
         TabIndex        =   39
         Top             =   720
         Width           =   1620
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Date      :"
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
         Left            =   8520
         TabIndex        =   32
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Maturity            :"
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
         TabIndex        =   29
         Top             =   4440
         Width           =   1620
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Date Release  :"
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
         TabIndex        =   28
         Top             =   3960
         Width           =   1650
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Amortization    :"
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
         TabIndex        =   27
         Top             =   3480
         Width           =   1635
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   26
         Top             =   3000
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Customer         :"
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
         TabIndex        =   25
         Top             =   2520
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Code                 :"
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
         TabIndex        =   24
         Top             =   2040
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Collector           :"
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
         TabIndex        =   23
         Top             =   1200
         Width           =   1650
      End
   End
   Begin VB.Label lblTotalPayment 
      Caption         =   "lblTotalPayment"
      Height          =   375
      Left            =   20400
      TabIndex        =   44
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "lblTime"
      Height          =   195
      Left            =   20400
      TabIndex        =   43
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUserlevel 
      AutoSize        =   -1  'True
      Caption         =   "lblUserlevel"
      Height          =   195
      Left            =   20400
      TabIndex        =   42
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   10680
      TabIndex        =   41
      Top             =   3480
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblBalance 
      AutoSize        =   -1  'True
      Caption         =   "lblBalance"
      Height          =   195
      Left            =   20400
      TabIndex        =   40
      Top             =   2520
      Width           =   735
   End
End
Attribute VB_Name = "frm_payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_payment
'    Project    : Project1
'
'    Description: [This procedure will provide Payments for Customer Loan]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Sub Compute()

        '<EhHeader>
        On Error GoTo Compute_Err

        TxtLog "Entered Compute"

        '</EhHeader>

        Dim TotalBalance As Double

        Dim Balance      As Double

100     TotalBalance = Val(txtTotaltAmortization.Text) - Val(txtPaymentmade.Text)
     
102     txtTotaltBalance.Text = Val(TotalBalance)

104     If Val(txtAmountPaid.Text) > Val(txtTotaltBalance.Text) Then
            
106         Balance = 0
108         lblPaymentsMade.Caption = txtAmountPaid.Text
            'lblOver.Caption = "0"
110         txtNewBalance.Text = "0"
        Else
112         Balance = Val(txtTotalBalance.Text) - Val(txtAmountPaid.Text)
114         lblPaymentsMade.Caption = txtAmountPaid.Text
            'lblOver.Caption = "0"
        End If

116     txtNewBalance.Text = Val(Balance)
            
118     If Val(txtNewBalance.Text) < 0 Then
            '  lblOver.Caption = txtTotaltBalance.Text
            
120         txtNewBalance.Text = "0"
        End If

        '<EhFooter>

        TxtLog "Exited Compute"

        Exit Sub

Compute_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.Compute", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub length()

        '<EhHeader>
        On Error GoTo length_Err

        TxtLog "Entered length"

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

        TxtLog "Exited length"

        Exit Sub

length_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.length", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub autoNumber()

        '<EhHeader>
        On Error GoTo autoNumber_Err

        TxtLog "Entered autoNumber"

        '</EhHeader>

100     If rsPayment1.State = 1 Then rsPayment1.Close
102     rsPayment1.Open "Select Max(ID) from tblPayment"
104     txtOR.Text = "1" & rsPayment1(0)

        'Benjamin Sumilhig And Jun rey Tavera 3-13-2015
        'Di man gyud na mazero nang record Count bisan zero ang value sa rsPayment kay naa man gyud na usa ka record count
        'mao dle modisplay ang OR1
        'if rsPayment.RecordCount = 0 then
        'textOR.Text = "OR1" Or txtOR.Text = "OR"
106     If txtOR.Text = "1" Then
108         txtOR.Text = "1"
        Else

            Dim newOR As Double

            'newOR = rsPayment1(0) + 1
            
            ' txtOR.Text = "" & rsPayment1(0) + 1
110         newOR = "" & rsPayment1(0) + 1
112         txtOR.Text = newOR
            
114         lblAuOR.Caption = txtOR.Text
        End If

        '<EhFooter>

        TxtLog "Exited autoNumber"

        Exit Sub

autoNumber_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.autoNumber", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub refresh_form_payment()

        '<EhHeader>
        On Error GoTo refresh_form_payment_Err

        TxtLog "Entered refresh_form_payment"

        '</EhHeader>

100     If rsPayment.State = 1 Then rsPayment.Close
102     rsPayment.Open "Select TOP 25 * from tblPayment Order By ID desc"
        'btnClose.Caption = "&Close"
        'btnAddpayment.Caption = "&Add Payment"
        'txtSearch.Enabled = True
104     txtCode.Text = ""
106     txtCollector.Text = ""
108     txtFCollector.Text = ""
110     txtOR.Text = ""
112     txtLoanID.Text = ""
114     txtCustomer.Text = ""
116     txtFirstName.Text = ""
118     txtPrincipal.Text = "0"
120     txtAmortization.Text = "0"
122     txtDateRelease.Text = ""
124     txtMaturity.Text = ""
126     txtTotaltAmortization.Text = "0"
128     txtPaymentmade.Text = "0"
130     txtTotaltBalance.Text = "0"
132     txtTotalBalance.Text = "0"
134     txtAmountPaid.Text = "0"
136     txtNewBalance.Text = "0"
        'txtCode.Enabled = False
        'txtAmountPaid.Enabled = False
138     txtOR.Text = lblAuOR.Caption

140     txtSearch.Enabled = False
142     btnAddpayment.Caption = "&Save Payment"
144     btnClose.Caption = "&Cancel"
146     txtCode.Enabled = True
148     txtAmountPaid.Enabled = True
150     txtCode.SetFocus
152     dtDate.Enabled = True

154     txtSearch.Enabled = False
        'btnAddpayment.Caption = "&Save Payment"
156     btnClose.Caption = "&Cancel"

158     Call autoNumber
        
160     Set DataGrid1.DataSource = rsPayment

        '<EhFooter>

        TxtLog "Exited refresh_form_payment"

        Exit Sub

refresh_form_payment_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.refresh_form_payment", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnAddpayment_Click()

        '<EhHeader>
        On Error GoTo btnAddpayment_Click_Err

        TxtLog "Entered btnAddpayment_Click"

        '</EhHeader>
    
100     Call autoNumber

102     If btnAddpayment.Caption = "&Add Payment" Then
104         txtSearch.Enabled = False
106         btnAddpayment.Caption = "&Save Payment"
108         btnClose.Caption = "&Cancel"
110         txtCode.Enabled = True
112         txtAmountPaid.Enabled = True
114         txtCode.SetFocus
116         dtDate.Enabled = True
118     ElseIf btnAddpayment.Caption = "&Save Payment" Then
120         Call Compute

122         If txtCode.Text = "" Or txtCode.Text = " " Then
124             MsgBox "Code should not be blank.", vbInformation
126             txtCode.Text = ""
128             txtCode.SetFocus
130         ElseIf txtCustomer.Text = "" Then
132             MsgBox "Please choose a customer first.", vbInformation
134             txtCode.Text = ""
136             txtCode.SetFocus
138         ElseIf DTPicker2.Value > dtDate.Value Then
140             MsgBox "Datepayment must not be later than date released. ", _
                        vbInformation, "Webplus Lending Corporation"
142             dtDate.SetFocus
144         ElseIf txtCustomer.Text = "" Or txtCollector.Text = "" Then
146             MsgBox "No Record Found", vbInformation, "Webplus Lending Corporation"
148             txtCode.Text = ""
150             txtCode.SetFocus
                   
            Else

                Dim LCode As String

152             LCode = ""
154             LCode = lblCuCode.Caption
                'This should include the FirstName, LastName, and Code of the Collector.
            
                'If rsPayment.State = 1 Then rsPayment.Close
                'rsPayment.Open "Select * from tblPayment Order By ORnumber desc"

                '  rsPayment.MoveFirst
156             If MsgBox("Are you sure you want to add a Payment for this customer?", _
                        vbQuestion + vbYesNo, "Webplus Lending Corporation") = vbYes _
                        Then
                    'ako ge comment
                    'Call autonumber

                    Dim f As Integer

158                 f = 0

160                 If rsPayment.State = 1 Then rsPayment.Close
162                 rsPayment.Open "Select * from tblPayment where Date = #" & _
                            dtDate.Value & "# and LoanID = " & txtLoanID.Text & _
                            " and Code = " & lblCuCode.Caption & " and Status = 'Good'"
                        
                    'If first time payment
164                 If rsPayment.RecordCount = 0 Then
                        'Present payment added to previous balance
166                     lblTotalPayment.Caption = Val(lblTotalPayment.Caption) + Val( _
                                lblPaymentsMade.Caption)
                        'ako Ge comment
                        'Call autonumber

168                     With rsPayment
170                         .AddNew
172                         !ORnumber = txtOR.Text
174                         !LoanID = txtLoanID.Text
176                         !Date = dtDate.Value
178                         !DateEncoded = txtTodayDate.Text
180                         !Collector = txtCollector.Text
182                         !code = lblCuCode.Caption
184                         LCode = lblCuCode.Caption
186                         !Customer = txtCustomer.Text
188                         !principal = txtPrincipal.Text
190                         !DateRelease = txtDateRelease.Text
If rsLoan!LOANTYPE <> 2 Then
192                         !Maturity = txtMaturity.Text
End If
194                         !Status = "Good"
196                         !Amortization = Val(txtTotaltAmortization.Text)
198                         !paymentsMade = Val(txtAmountPaid.Text)
200                         !TotalBalance = Val(txtTotalBalance.Text)
202                         !NewBalance = Val(txtNewBalance.Text)
204                         !User = lblUser.Caption
                            '!Over = Val(lblOver.Caption)
206                         !TotalPayment = Val(lblTotalPayment.Caption)
208                         !CollectorCode = lblCollectorCode.Caption
210                         !CollectorFname = lblCollectorFname.Caption
212                         .Update
                        End With

214                     f = 1
                        'Updates the balance of customer

                        'Call auditPayment(rsLoan!LoanID)
                        'Call refresh_form_payment
                        'frm_payment.lblUser.Caption = MDIForm1.lblUserName.Caption
                        'frm_payment.lblUserlevel.Caption = MDIForm1.lblUserlevel.Caption
                        'Me.Show
                        'txtSearch.Enabled = True
                        'Kung ika duha or more than one na payment
216                 ElseIf rsPayment.RecordCount <> 0 Then

                        ' Kung mag payment same date
218                     If MsgBox( _
                                "This customer has already paid on this day. Do you still want to save another payment for this Customer?", _
                                vbQuestion + vbYesNo, "Webplus Lending Corporation") = _
                                vbYes Then
                            'Present payment added to previous balance
220                         lblTotalPayment.Caption = Val(lblTotalPayment.Caption) + _
                                    Val(lblPaymentsMade.Caption)
                            'Call Compute
                            'Call autonumber

222                         With rsPayment
224                             .AddNew
226                             !ORnumber = txtOR.Text
228                             !LoanID = txtLoanID.Text
230                             !Date = dtDate.Value
232                             !DateEncoded = txtTodayDate.Text
234                             !Collector = txtCollector.Text
236                             !code = lblCuCode.Caption
238                             LCode = lblCuCode.Caption
240                             !Customer = txtCustomer.Text
242                             !principal = txtPrincipal.Text
244                             !DateRelease = txtDateRelease.Text
If rsLoan!LOANTYPE = 1 Then
246                             !Maturity = txtMaturity.Text
Else
End If
248                             !Status = "Good"
250                             !Amortization = Val(txtTotaltAmortization.Text)
252                             !paymentsMade = Val(lblPaymentsMade.Caption)
254                             !TotalBalance = Val(txtTotalBalance.Text)
256                             !NewBalance = Val(txtNewBalance.Text)
258                             !User = lblUser.Caption
                                '!Over = Val(lblOver.Caption)
260                             !TotalPayment = Val(lblTotalPayment.Caption)
262                             !CollectorCode = lblCollectorCode.Caption
264                             !CollectorFname = lblCollectorFname.Caption
266                             .Update
                            End With

268                         f = 1
                           
                            'Call auditPayment(rsLoan!LoanID)

270                         If rsTrail.State = 1 Then rsTrail.Close
272                         rsTrail.Open "Select * from tblTrail "

                            'Records new activity of User
274                         With rsTrail
276                             .AddNew
278                             !UserName = lblUser.Caption
280                             !userlevel = lblUserlevel.Caption
282                             !Activity = "Add New Payment"
284                             !Time = lblTime.Caption
286                             !Date = txtTodayDate.Text
288                             .Update
                            End With

                            'Audit Payment

290                         txtSearch.Enabled = True
                        End If
                    End If

                Else
292                 txtSearch.Enabled = False

                End If
                
294             If f = 1 Then

                    'Updates the balance of customer
296                 If rsCustomer.State = 1 Then rsCustomer.Close
298                 rsCustomer.Open "Select * from tblCustomer where Code = " & LCode & ""
                    If rsLoan!LOANTYPE = 1 Then
300                     rsCustomer!Balance = txtNewBalance.Text
                    Else
                        rsCustomer!EMERGENCYBalance = txtNewBalance.Text
                    End If
302                 rsCustomer.Update

                    'Updates the Status of Loan Record of the customer , if Full Paid
304                 If rsLoan.State = 1 Then rsLoan.Close
306                 rsLoan.Open "Select * from tblLoan where code = " & LCode & _
                            "and Status = '" & "Good" & "' "

308                 If txtNewBalance.Text = 0 Then
310                     rsLoan!Status = "Full Paid"
312                     rsLoan!TotalAmortization = txtNewBalance.Text
314                     rsLoan!TotalPayment = lblTotalPayment.Caption
316                     rsLoan.Update
                
318                     rsPayment!Status = "Full Paid"
320                     rsPayment!User = lblUser.Caption
322                     rsPayment.Update
324                     MsgBox "Customer is now Full Paid", vbInformation, _
                                "Webplus Lending Corporation"
                    Else
                        'If not Full paid
326                     rsLoan!TotalAmortization = txtNewBalance.Text
328                     rsLoan!TotalPayment = lblTotalPayment.Caption
330                     rsLoan.Update
                    End If
                            
                    'Updates the collection of collector
                    Dim collection As Double

332                 If rsCollData2.State = 1 Then rsCollData2.Close
334                 rsCollData2.Open "Select * from tblColl_Data where Code = " & _
                            rsPayment!CollectorCode & ""

336                 If rsCollData2.RecordCount = 0 Then
                    Else
                        ' rsCollData2!collection = Val(rsCollData2!collection) + Val(txtAmountPaid.Text)
338                     collection = Val(rsCollData2!collection) + Val(txtAmountPaid.Text)
                        'rsCollData2!YTDCollection = Val(rsCollData2!YTDCollection) + collection
340                     rsCollData2!YTDCollection = collection
342                     rsCollData2.Update
                    End If

                    'Call auditPayment(rsLoan!LoanID)
344                 Call refresh_form_payment
346                 frm_payment.lblUser.Caption = MDIForm1.lblusername.Caption
348                 frm_payment.lblUserlevel.Caption = MDIForm1.lblUserlevel.Caption
                    'Me.Show
350                 txtSearch.Enabled = True
                End If
                            
            End If
        End If

        'Call setgrid

        '<EhFooter>

        TxtLog "Exited btnAddpayment_Click"

        Exit Sub

btnAddpayment_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.btnAddpayment_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnClose_Click()

        '<EhHeader>
        On Error GoTo btnClose_Click_Err

        TxtLog "Entered btnClose_Click"

        '</EhHeader>
        
100     Call autoNumber
        
102     If btnClose.Caption = "&Close" Then
104         Unload Me
        Else
106         btnClose.Caption = "&Close"
108         btnAddpayment.Caption = "&Add Payment"
110         txtSearch.Enabled = True
112         txtCode.Text = ""
114         txtCollector.Text = ""
116         txtFCollector.Text = ""
118         txtOR.Text = ""
120         txtLoanID.Text = ""
122         txtCustomer.Text = ""
124         txtFirstName.Text = ""
126         txtPrincipal.Text = "0"
128         txtAmortization.Text = "0"
130         txtDateRelease.Text = ""
132         txtMaturity.Text = ""
134         txtTotaltAmortization.Text = "0"
136         txtPaymentmade.Text = "0"
138         txtTotaltBalance.Text = "0"
140         txtTotalBalance.Text = "0"
142         txtAmountPaid.Text = "0"
144         txtNewBalance.Text = "0"
146         txtCode.Enabled = False
148         txtAmountPaid.Enabled = False
150         txtOR.Text = lblAuOR.Caption
152         dtDate.Enabled = False
            
            ' Call autonumber
154         If rsPayment.State = 1 Then rsPayment.Close
156         rsPayment.Open "Select TOP 25 * from tblPayment Order By ORnumber desc"
158         Set DataGrid1.DataSource = rsPayment
160         DataGrid1.Width = Me.Width
162         DataGrid1.Columns(0).Width = 0
164         DataGrid1.Columns(1).Width = 1100
166         DataGrid1.Columns(3).Width = 900
168         DataGrid1.Columns(4).Width = 900
            '172         DataGrid1.Columns(6).Width = 1200
            '174         DataGrid1.Columns(7).Width = 1000
            '176         DataGrid1.Columns(8).Width = 1000
            '178         DataGrid1.Columns(9).Width = 1100
            '180         DataGrid1.Columns(10).Width = 1100
            '182         DataGrid1.Columns(12).Width = 1100
            '184         DataGrid1.Columns(15).Width = 1100
        End If

170     MDIForm1.Picture1.Visible = True

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.btnClose_Click", Erl

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
        ErrReport Err.Description, "LendingClient.frm_payment.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub dtDate_Change()

        '<EhHeader>
        On Error GoTo dtDate_Change_Err

        TxtLog "Entered dtDate_Change"

        '</EhHeader>

100     If rsCollData.State = 1 Then rsCollData.Close
102     rsCollData.Open "Select * from tblColl_Data where Code = " & _
                lblCollectorCode.Caption & " and  DateEmployed <= #" & dtDate.Value & _
                "# order by DateEmployed DESC"
                
104     If rsCollData.RecordCount <> 0 Then
106         rsCollData.MoveFirst
108         txtCollector.Text = rsCollData!lastname
110         lblCollectorFname.Caption = rsCollData!firstname
112         txtFCollector.Text = rsCollData!firstname
        Else
114         lblCollectorFname.Caption = ""
116         txtCollector.Text = ""
118         txtFCollector.Text = ""
        End If

        '<EhFooter>

        TxtLog "Exited dtDate_Change"

        Exit Sub

dtDate_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.dtDate_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub dtDate_Click()

        '<EhHeader>
        On Error GoTo dtDate_Click_Err

        TxtLog "Entered dtDate_Click"

        '</EhHeader>

100     If rsCollData.State = 1 Then rsCollData.Close
102     rsCollData.Open "Select * from tblColl_Data where Code = '" & _
                lblCollectorCode.Caption & "' and  DateEmployed <= #" & dtDate.Value & _
                "# order by DateEmployed DESC"
                
104     If rsCollData.RecordCount <> 0 Then
106         rsCollData.MoveFirst
108         txtCollector.Text = rsCollData!lastname
110         lblCollectorFname.Caption = rsCollData!firstname
112         txtFCollector.Text = rsCollData!firstname
        Else
114         lblCollectorFname.Caption = ""
116         txtCollector.Text = ""
118         txtFCollector.Text = ""
        End If

        '<EhFooter>

        TxtLog "Exited dtDate_Click"

        Exit Sub

dtDate_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.dtDate_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

        'hide the  unnecessary label

100     lblBalance.Visible = False
102     lblTime.Visible = False
104     lblUserlevel.Visible = False
106     lblTotalPayment.Visible = False
108     Label_NewOr.Visible = False
110     lblCollectorCode.Visible = False
112     lblOver.Visible = False

        'Benjamin sumilhig

114     Call connect
116     Call Loan
118     Call Customer
120     Call Collector
122     Call payment
124     Call Trail
126     Call CollData
128     Call payment1
130     Me.Show
132     Me.SetFocus
134     Me.WindowState = vbMaximized

136     If rsPayment.State = 1 Then rsPayment.Close
138     rsPayment.Open "Select TOP 25 * from tblPayment where Code <> -1 Order By ID desc"
140     Call autoNumber
        
142     Set DataGrid1.DataSource = rsPayment
        
144     DataGrid1.Width = Me.Width
146     DataGrid1.Columns(0).Width = 0
148     DataGrid1.Columns(1).Width = 1100
150     DataGrid1.Columns(3).Width = 900
152     DataGrid1.Columns(4).Width = 900
        '132     DataGrid1.Columns(6).Width = 1200
        '134     DataGrid1.Columns(7).Width = 1000
        '136     DataGrid1.Columns(8).Width = 1000
        '138     DataGrid1.Columns(9).Width = 1100
        '140     DataGrid1.Columns(10).Width = 1100
        '142     DataGrid1.Columns(12).Width = 1100
        '144     DataGrid1.Columns(15).Width = 1100
        '
        '150     With Me
        '152         .Top = (Screen.Height - .Height) / 2
        '154         .Left = (Screen.Width - .Width) / 2
        '        End With

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    txtTodayDate.Text = Date
    lblTime.Caption = Time
    dtDate.Value = Date
    
    Timer1.Enabled = False
    
End Sub

Private Sub txtAmountPaid_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtAmountPaid_KeyPress_Err

        TxtLog "Entered txtAmountPaid_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321"

        '        'Check for Only One Decimal Point In Textbox
        '106     If KeyAscii = 46 Then
        '
        '            'If more than one decimal point is typed, only one decimal will be printed
        '108         If InStr(1, txtAmountPaid.Text, ".") > 0 Then
        '110             KeyAscii = 0
        '112             MsgBox "Multiple decimal points are not allowed.", vbOKOnly + vbInformation
        '
        '                Exit Sub
        '
        '            End If

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8

108     ElseIf KeyAscii = 13 Then

            'txtTotaltAmortization.Text = txtAmortization.Text
110         Call Compute
            '   btnAddpayment.SetFocus
            
112         Call btnAddpayment_Click
            
        Else
114         KeyAscii = 0
        
        End If

        '<EhFooter>

        TxtLog "Exited txtAmountPaid_KeyPress"

        Exit Sub

txtAmountPaid_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.txtAmountPaid_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtAmountPaid_LostFocus()

        '<EhHeader>
        On Error GoTo txtAmountPaid_LostFocus_Err

        TxtLog "Entered txtAmountPaid_LostFocus"

        '</EhHeader>

100     txtTotaltAmortization.Text = txtAmortization.Text
102     Call Compute

        '<EhFooter>

        TxtLog "Exited txtAmountPaid_LostFocus"

        Exit Sub

txtAmountPaid_LostFocus_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.txtAmountPaid_LostFocus", _
                Erl

        Resume Next

        '</EhFooter>

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

110         Call length
            
            'Searching for existing Loan based on Code
112         If rsLoan.State = 1 Then rsLoan.Close
114         rsLoan.Open "Select * from tblLoan where Code = " & txtCode.Text & _
                    " and Status <> '" & "Full Paid" & "' and Status = '" & "Good" & _
                    "' and TotalAmortization > 0 order by LoanID DESC"

            'If loan doesn't exist
116         If rsLoan.RecordCount = 0 Then
118             MsgBox "No record found.", vbInformation
120             txtCode.Text = ""
122             txtCode.SetFocus

                Exit Sub

124         ElseIf rsLoan!Status = "FullPaid" Then
126             MsgBox "Customer balance is Already Full Paid.", vbInformation
128             txtCode.SetFocus
                'If there is new loan
130         ElseIf rsLoan.RecordCount = 1 Then

132             With rsLoan
134                 txtLoanID.Text = !LoanID
                    'txtCollector.Text = !Collector
                    'txtFCollector.Text = !CollectorFname
136                 txtCustomer.Text = !Customer
138                 txtFirstName.Text = !firstname
140                 txtPrincipal.Text = !principal
                    
142                 txtAmortization.Text = !total
144                 txtTotaltAmortization.Text = !total
146                 txtDateRelease.Text = !DateRelease
148                 DTPicker2.Value = !DateRelease
If !LOANTYPE <> 2 Then
150                 txtMaturity.Text = !Maturity
End If
152                 txtTotalBalance.Text = !TotalAmortization
154                 lblTotalPayment.Caption = !TotalPayment
156                 lblCollectorCode.Caption = !CollectorCode
                    'lblCollectorFname.Caption = !CollectorFname
158                 lblCuCode.Caption = !code
                End With

160             If rsCollData.State = 1 Then rsCollData.Close
162             rsCollData.Open "Select * from tblColl_Data where Code = " & _
                        lblCollectorCode.Caption & " and  DateEmployed <= #" & _
                        dtDate.Value & "# order by DateEmployed DESC"
                
164             If rsCollData.RecordCount <> 0 Then
166                 rsCollData.MoveFirst
168                 txtCollector.Text = rsCollData!lastname
170                 lblCollectorFname.Caption = rsCollData!firstname
172                 txtFCollector.Text = rsCollData!firstname
                Else
174                 lblCollectorFname.Caption = ""
176                 txtCollector.Text = ""
178                 txtFCollector.Text = ""
                End If

180             txtAmountPaid.SetFocus

182             lblBalance.Caption = rsCustomer!Balance
184             txtPaymentmade.Text = lblTotalPayment.Caption
186             Call Compute
                'For the renewal loan
188         ElseIf rsLoan.RecordCount > 1 Then

                ' If MsgBox("Are you sure you want to add this payment?", vbQuestion + vbYesNo, "Jajavi Lending Corporation") = vbYes Then

190             If rsLoan.State = 1 Then rsLoan.Close
192             rsLoan.Open "Select * from tblLoan where Code = " & txtCode.Text & _
                        " and Status <> '" & "Full Paid" & "' "

194             If rsLoan.RecordCount <> 0 Then
196                 rsLoan.MoveFirst

198                 With rsLoan
200                     txtLoanID.Text = !LoanID
202                     txtCollector.Text = !Collector
204                     txtFCollector.Text = !CollectorFname
206                     txtCustomer.Text = !Customer
208                     txtFirstName.Text = !firstname
210                     txtPrincipal.Text = !principal
                    
212                     txtAmortization.Text = !total
214                     txtTotaltAmortization.Text = !total
216                     txtDateRelease.Text = !DateRelease
218                     DTPicker2.Value = !DateRelease
220                     txtMaturity.Text = !Maturity
222                     txtTotalBalance.Text = !TotalAmortization
224                     lblTotalPayment.Caption = !TotalPayment
226                     lblCollectorCode.Caption = !CollectorCode
228                     lblCollectorFname.Caption = !CollectorFname
230                     lblCuCode.Caption = !code
                    End With
                        
232                 If rsCollData.State = 1 Then rsCollData.Close
234                 rsCollData.Open "Select * from tblColl_Data where Code = " & _
                            lblCollectorCode.Caption & " and  DateEmployed <= #" & _
                            dtDate.Value & "# order by DateEmployed DESC"
                
236                 If rsCollData.RecordCount <> 0 Then
238                     rsCollData.MoveFirst
240                     txtCollector.Text = rsCollData!lastname
242                     lblCollectorFname.Caption = rsCollData!firstname
244                     txtFCollector.Text = rsCollData!firstname
                    Else
246                     lblCollectorFname.Caption = ""
248                     txtCollector.Text = ""
250                     txtFCollector.Text = ""
                    End If
                        
                End If

252             txtAmountPaid.SetFocus
            Else
254             KeyAscii = 0
            
            End If

        Else
256         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtCode_KeyPress"

        Exit Sub

txtCode_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.txtCode_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsPayment.State = 1 Then rsPayment.Close

102     rsPayment.Open "Select * from tblPayment where ((Customer like '" & _
                txtSearch.Text & "%') and (Customer <> 'Over')) or ORnumber like '" & _
                Val(txtSearch.Text) & "%' or ((Code like '%" & Val(txtSearch.Text) & _
                "%') and (Code <> -1)) or Date like '" & txtSearch.Text & _
                "%' or DateEncoded like '" & txtSearch.Text & "%' Order By id desc"
        
104     Set DataGrid1.DataSource = rsPayment

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_payment.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

