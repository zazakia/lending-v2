VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_PaymentReverse 
   Caption         =   "frmPaymentReverse"
   ClientHeight    =   10725
   ClientLeft      =   1425
   ClientTop       =   1620
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   ScaleHeight     =   10725
   ScaleWidth      =   20370
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Payment Reverse Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20295
      Begin VB.TextBox txtFCollector 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   55
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtNewOR 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   52
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtLoanID 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   47
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   2640
         TabIndex        =   43
         Top             =   5160
         Width           =   3375
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3735
         Left            =   240
         TabIndex        =   42
         Top             =   5760
         Width           =   19935
         _ExtentX        =   35163
         _ExtentY        =   6588
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
      Begin VB.TextBox txtR 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   1800
         TabIndex        =   38
         Text            =   "OR"
         Top             =   720
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   7920
         Top             =   3000
      End
      Begin VB.TextBox txtTodayDate 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   34
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   2760
         TabIndex        =   32
         Top             =   9600
         Width           =   2055
      End
      Begin VB.CommandButton btnReverse 
         Caption         =   "&Reverse Payment"
         Height          =   615
         Left            =   120
         TabIndex        =   31
         Top             =   9600
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Section"
         Height          =   2895
         Left            =   10680
         TabIndex        =   13
         Top             =   2760
         Width           =   6255
         Begin VB.TextBox txtDatePayment 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2640
            TabIndex        =   48
            Top             =   2280
            Width           =   2775
         End
         Begin VB.TextBox txtNBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2640
            TabIndex        =   16
            Text            =   "0"
            Top             =   1560
            Width           =   2775
         End
         Begin VB.TextBox txtAmountPaid 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2640
            TabIndex        =   15
            Text            =   "0"
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox txtOutBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2640
            TabIndex        =   14
            Text            =   "0"
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Paid:"
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
            TabIndex        =   49
            Top             =   2280
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Outstanding Balance  :"
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
            TabIndex        =   19
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Amount paid                 :"
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
            TabIndex        =   18
            Top             =   960
            Width           =   2430
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Balance               :"
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
            TabIndex        =   17
            Top             =   1560
            Width           =   2415
         End
         Begin VB.Line Line5 
            X1              =   2520
            X2              =   5400
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line4 
            X1              =   2520
            X2              =   5400
            Y1              =   1440
            Y2              =   1440
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Balance Info"
         Height          =   2295
         Left            =   10680
         TabIndex        =   9
         Top             =   360
         Width           =   5895
         Begin VB.TextBox txtTotaltBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2040
            TabIndex        =   12
            Text            =   "0"
            Top             =   1320
            Width           =   3735
         End
         Begin VB.TextBox txtPaymentmade 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2040
            TabIndex        =   11
            Text            =   "0"
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtTotaltAmortization 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2040
            TabIndex        =   10
            Text            =   "0"
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Total Balance     :"
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
            TabIndex        =   30
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Payment              :"
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
            TabIndex        =   29
            Top             =   720
            Width           =   1830
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Balance               :"
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
            TabIndex        =   28
            Top             =   240
            Width           =   1830
         End
         Begin VB.Line Line3 
            X1              =   1800
            X2              =   5760
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line2 
            X1              =   1800
            X2              =   5760
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   5760
            Y1              =   1200
            Y2              =   1200
         End
      End
      Begin VB.TextBox txtMaturity 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txtDateRelease 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox txtAmortization 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Text            =   "0"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox txtPrincipal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Text            =   "0"
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtCustomer 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtCollector 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtOR 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblTotalBalance 
         Caption         =   "o"
         Height          =   255
         Left            =   8760
         TabIndex        =   61
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblPaOR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblPaOR"
         Height          =   195
         Left            =   5760
         TabIndex        =   60
         Top             =   3720
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label terwe 
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
         Left            =   2640
         TabIndex        =   59
         Top             =   4680
         Width           =   90
      End
      Begin VB.Label lblLabel18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH USING CUSTOMER, LOANID OR DATE"
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
         Left            =   2760
         TabIndex        =   58
         Top             =   4680
         Width           =   4485
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "count"
         Height          =   195
         Left            =   5400
         TabIndex        =   57
         Top             =   2880
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblCollectorCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   10080
         TabIndex        =   56
         Top             =   1080
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblTotalPayment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
         Height          =   195
         Left            =   16680
         TabIndex        =   54
         Top             =   1560
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblLabel17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New ORNumber  :"
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
         Left            =   4560
         TabIndex        =   53
         Top             =   720
         Width           =   1905
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
         Left            =   1080
         TabIndex        =   51
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblNoteOR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note: OR Number must be typed by numbers only"
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
         Left            =   1320
         TabIndex        =   50
         Top             =   360
         Width           =   4410
      End
      Begin VB.Label lblLoanID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LoanID:"
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
         Left            =   3840
         TabIndex        =   46
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblSearchHere 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Here      :"
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
         Left            =   480
         TabIndex        =   44
         Top             =   5160
         Width           =   1770
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
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
         TabIndex        =   33
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Maturity          :"
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
         Top             =   4080
         Width           =   1500
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Date Release:"
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
         Top             =   3600
         Width           =   1530
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Amortization  :"
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
         Top             =   3120
         Width           =   1515
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Principal         :"
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
         Top             =   2640
         Width           =   1485
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Customer       :"
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
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Code               :"
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
         TabIndex        =   22
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Collector         : "
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
         Top             =   1200
         Width           =   1590
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "OR Number   : "
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
         TabIndex        =   20
         Top             =   720
         Width           =   1560
      End
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   195
      Left            =   12840
      TabIndex        =   45
      Top             =   6840
      Width           =   45
   End
   Begin VB.Label lblPaymentMade 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   11400
      TabIndex        =   41
      Top             =   1800
      Width           =   45
   End
   Begin VB.Label lblminus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   195
      Left            =   11400
      TabIndex        =   40
      Top             =   4800
      Width           =   45
   End
   Begin VB.Label lblOR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   11280
      TabIndex        =   39
      Top             =   5280
      Width           =   45
   End
   Begin VB.Label lblTime 
      Height          =   135
      Left            =   20280
      TabIndex        =   37
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label lblUserlevel 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   12840
      TabIndex        =   36
      Top             =   4320
      Width           =   45
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "Label16"
      Height          =   195
      Left            =   20280
      TabIndex        =   35
      Top             =   3120
      Width           =   570
   End
End
Attribute VB_Name = "frm_PaymentReverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_PaymentReverse
'    Project    : Project1
'
'    Description: [This Module will reverse the existing payment of the customer.]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Sub autoNumber()

        '<EhHeader>
        On Error GoTo autoNumber_Err

        TxtLog "Entered autoNumber"

        '</EhHeader>

        'If rsPayment.State = 1 Then rsPayment.Close
        'rsPayment.Open " Select * from tblPayment order by ORnumber desc"
        'If rsPayment.RecordCount = 0 Then
        '         txtNewOR.Text = "OR1"
        '        Else
        '    rsPayment.MoveLast
        '         txtNewOR.Text = "OR" & Format(Right(rsPayment!ORnumber, 5) + 1, "00000")
        '        End If

100     Call payment1

102     If rsPayment1.State = 1 Then rsPayment1.Close
104     rsPayment1.Open "Select Max(ID) from tblPayment"
106     txtNewOR.Text = "OR" & rsPayment1(0)

108     If txtNewOR.Text = "" Then
110         txtNewOR.Text = ""
112         btnReverse.Enabled = False
        Else

            Dim newOR As Double

114         newOR = rsPayment1(0) + 1
116         txtNewOR.Text = newOR
            ' txtNewOR.Text = "OR" & rsPayment1(0) + 1

            'lblAuOR.Caption = txtOR.Text
        End If

        '<EhFooter>

        TxtLog "Exited autoNumber"

        Exit Sub

autoNumber_Err:
        ErrReport Err.Description, "LendingClient.frm_PaymentReverse.autoNumber", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub setgrid2()

        '<EhHeader>
        On Error GoTo setgrid2_Err

        TxtLog "Entered setgrid2"

        '</EhHeader>

100     If rsPayment.State = 1 Then rsPayment.Close
102     rsPayment.Open "Select TOP 25 * from tblPayment Order By ID desc"
104     Set DataGrid1.DataSource = rsPayment
106     DataGrid1.Width = Me.Width
108     DataGrid1.Columns(0).Width = 0
110     DataGrid1.Columns(1).Width = 1100
112     DataGrid1.Columns(3).Width = 900
114     DataGrid1.Columns(4).Width = 900
116     DataGrid1.Columns(6).Width = 1200
118     DataGrid1.Columns(7).Width = 1000
120     DataGrid1.Columns(8).Width = 1000
122     DataGrid1.Columns(9).Width = 1100
124     DataGrid1.Columns(10).Width = 1100
126     DataGrid1.Columns(12).Width = 1100
128     DataGrid1.Columns(15).Width = 1100

        '<EhFooter>

        TxtLog "Exited setgrid2"

        Exit Sub

setgrid2_Err:
        ErrReport Err.Description, "LendingClient.frm_PaymentReverse.setgrid2", Erl

        Resume Next

        '</EhFooter>

End Sub

'by BenJun 03-11-2015 10:00PM
'mao neh ang mo subract sa tblCOll_data!COlletion and tblColl_Data!YTDCollection

Public Sub Subtract_Coll(x As Double)

        '<EhHeader>
        On Error GoTo Subtract_Coll_Err

        TxtLog "Entered Subtract_Coll"

        '</EhHeader>

100     If rsCollData2.State = 1 Then rsCollData2.Close
102     rsCollData2.Open "Select * from tblColl_Data where Code = " & x & ""

104     If rsCollData2.RecordCount = 0 Then
        Else
106         rsCollData2!collection = Val(rsCollData2!collection) - Val(txtAmountPaid.Text)
108         rsCollData2!YTDCollection = Val(rsCollData2!YTDCollection) - Val( _
                    txtAmountPaid.Text)
110         rsCollData2.Update
        End If

        '<EhFooter>

        TxtLog "Exited Subtract_Coll"

        Exit Sub

Subtract_Coll_Err:

        ErrReport Err.Description, "LendingClient.frm_PaymentReverse.Subtract_Coll", Erl

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
106         txtOR.Enabled = False
108         btnClose.Caption = "&Close"
110         btnReverse.Caption = "&Reverse Payment"
112         txtOR.Text = ""
114         txtCollector.Text = ""
116         txtFCollector.Text = ""
118         txtCode.Text = ""
120         txtCustomer.Text = ""
122         txtPrincipal.Text = "0"
124         txtAmortization.Text = "0"
126         txtDateRelease.Text = ""
128         txtMaturity.Text = ""
130         txtLoanID.Text = ""
132         txtTotaltAmortization.Text = "0"
134         txtPaymentmade.Text = "0"
136         txtTotaltBalance.Text = "0"
138         txtOutBalance.Text = "0"
140         txtAmountPaid.Text = "0"
142         txtNBalance.Text = "0"
144         txtDatePayment.Text = ""

146         If rsPayment.State = 1 Then rsPayment.Close
148         rsPayment.Open "Select TOP 25 * from tblPayment Order By ORnumber desc"
150         Set DataGrid1.DataSource = rsPayment
152         DataGrid1.Width = Me.Width
154         DataGrid1.Columns(0).Width = 0
156         DataGrid1.Columns(1).Width = 1100
158         DataGrid1.Columns(3).Width = 900
160         DataGrid1.Columns(4).Width = 900
162         DataGrid1.Columns(6).Width = 1200
164         DataGrid1.Columns(7).Width = 1000
166         DataGrid1.Columns(8).Width = 1000
168         DataGrid1.Columns(9).Width = 1100
170         DataGrid1.Columns(10).Width = 1100
172         DataGrid1.Columns(12).Width = 1100
174         DataGrid1.Columns(15).Width = 1100
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_PaymentReverse.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnReverse_Click()

        '<EhHeader>
        On Error GoTo btnReverse_Click_Err

        TxtLog "Entered btnReverse_Click"

        '</EhHeader>
    
        Dim LoanID      As Long

        'Begin: benjack 3-1-15
        Dim txtcode_ben As String

100     txtcode_ben = txtCode.Text
        'End: benjack 3-1-15
102     LoanID = 0
    
104     txtSearch.Enabled = False
106     Call autoNumber

108     If rsPayment.State = 1 Then rsPayment.Close
110     rsPayment.Open "Select * from tblPayment Order By ORnumber desc"
112     Set DataGrid1.DataSource = rsPayment
114     Call setgrid2

116     If btnReverse.Caption = "&Reverse Payment" Then
118         btnReverse.Caption = "&Update"
120         btnClose.Caption = " &Cancel"
122         txtOR.Enabled = True
124         txtOR.SetFocus
126     ElseIf txtcode_ben = "" Then
128         MsgBox "Press enter key first before you can click the update button"
130         txtOR.SetFocus
132     ElseIf btnReverse.Caption = "&Update" Then

            Dim RPay As String

134         RPay = ""
136         RPay = lblPaOR.Caption

138         If rsPayment.State = 1 Then rsPayment.Close
140         rsPayment.Open "Select * from tblPayment where ORnumber = " & RPay & " "
        
            'Print rsPayment!Code

            'get the loanID of the active, current loan
142         If rsLoan1.State = 1 Then rsLoan1.Close
144         rsLoan1.Open "Select * from tblLoan where Code = " & rsPayment!code & _
                    " and (Status = 'Good' or Status = 'Full Paid') Order By LoanID desc"
        
146         If rsLoan1.RecordCount <> 0 Then
148             rsLoan1.MoveFirst
150             Print rsLoan1!LoanID
            End If
        
152         If txtCode.Text = "" Then
154             MsgBox "No record found.", vbInformation, "Webplus Lending Corporation"

156             If rsPayment.State = 1 Then rsPayment.Close
158             rsPayment.Open "Select * from tblPayment Order By ORnumber desc"
160             Set DataGrid1.DataSource = rsPayment
162             txtCode.Text = ""
        
                'pero what if wala syay past na loan?
164         ElseIf txtCode.Text = "Over" Then
166             MsgBox "Invalid transaction for reversing OR number"
        
168         ElseIf (rsLoan1!LoanID <> rsPayment!LoanID) Then
        
170             MsgBox _
                        "This payment is not of the customer's current/active loan. Please recheck."
        
            Else

172             If MsgBox( _
                        "Are you sure you want to Reverse this Payment for this customer?", _
                        vbQuestion + vbYesNo, "J Lending Corporation") = vbYes Then
                
                    'Dim test As String
            
174                 If rsPayment.State = 1 Then rsPayment.Close
176                 rsPayment.Open "Select * from tblPayment where ORnumber = " & _
                            lblPaOR.Caption & " "
178                 LoanID = rsPayment!LoanID
                
180                 If rsPayment.RecordCount <> 0 Then
                   
182                     If rsPayment!Status = "Full Paid" Then

184                         With rsPayment
                                ' test = lblPaOR.Caption
186                             !Status = "Reversed"
188                             Call autoNumber
190                             .AddNew
192                             !ORnumber = txtNewOR.Text
194                             !LoanID = txtLoanID.Text
196                             !Date = txtTodayDate.Text
198                             !Collector = txtCollector.Text
200                             !code = txtCode.Text
202                             !Customer = txtCustomer.Text
204                             !principal = txtPrincipal.Text
206                             !DateRelease = txtDateRelease.Text
208                             !Maturity = txtMaturity.Text
210                             !Amortization = txtTotaltAmortization.Text
212                             !paymentsMade = txtPaymentmade.Text
214                             !TotalBalance = txtTotaltBalance.Text
216                             !NewBalance = Val(lblTotalBalance.Caption)
218                             !TotalPayment = Val(lblTotalPayment.Caption) - Val( _
                                        txtAmountPaid.Text)
220                             !DateEncoded = txtTodayDate.Text
                                '!Over = "0"
222                             !CollectorFname = txtFCollector.Text
224                             !CollectorCode = lblCollectorCode.Caption
226                             !Status = "Reversing"
228                             !User = lblUser.Caption
230                             .Update
                            End With

                            'by benJun 3-11-2015 10:28 PM
                            'call subtract_Coll to excute sa pag minus sa tblCOll_Data collection ug YTD collection kung rspayment!status kay Good
232                         Call Subtract_Coll(rsPayment!CollectorCode)

234                     ElseIf rsPayment!Status = "Good" Then
                    
236                         With rsPayment
                                'test = lblPaOR.Caption
238                             !Status = "Reversed"
240                             Call autoNumber
242                             .AddNew
244                             !ORnumber = txtNewOR.Text
246                             !LoanID = txtLoanID.Text
248                             !Date = txtTodayDate.Text
250                             !Collector = txtCollector.Text
252                             !code = txtCode.Text
254                             !Customer = txtCustomer.Text
256                             !principal = txtPrincipal.Text
258                             !DateRelease = txtDateRelease.Text
260                             !Maturity = txtMaturity.Text
262                             !Amortization = txtTotaltAmortization.Text
264                             !paymentsMade = txtPaymentmade.Text
266                             !TotalBalance = txtTotaltBalance.Text
268                             !NewBalance = Val(txtTotaltBalance.Text)
270                             !TotalPayment = Val(lblTotalPayment.Caption) - Val( _
                                        txtAmountPaid.Text)
272                             !DateEncoded = txtTodayDate.Text
                                '!Over = "0"
274                             !CollectorFname = txtFCollector.Text
276                             !CollectorCode = lblCollectorCode.Caption
278                             !Status = "Reversing"
280                             !User = lblUser.Caption
282                             .Update
                            End With

                            'by benJun 3-11-2015 10:28 PM
                            'call subtract_Coll to excute sa pag minus sa tblCOll_Data collection ug YTD collection kung rspayment!status kay Full Paid
284                         Call Subtract_Coll(rsPayment!CollectorCode)

                        End If

286                     If rsLoan.State = 1 Then rsLoan.Close
288                     rsLoan.Open "Select * from tblLoan where LoanID = " & _
                                txtLoanID.Text & ""
290                     rsLoan!TotalAmortization = rsPayment!NewBalance
292                     rsLoan!Status = "Good"
294                     rsLoan!TotalPayment = Val(lblTotalPayment.Caption) - Val( _
                                txtAmountPaid.Text)
296                     rsLoan.Update
                    
298                     If rsCustomer.State = 1 Then rsCustomer.Close
300                     rsCustomer.Open "Select * from tblCustomer where Code = " & _
                                txtCode.Text & ""
302                     rsCustomer!Balance = rsPayment!NewBalance
                    
304                     If rsPayment.State = 1 Then rsPayment.Close
306                     rsPayment.Open "Select * from tblPayment where LoanID = " & _
                                txtLoanID.Text & " and Status = 'Full Paid' "

308                     If rsPayment.RecordCount <> 0 Then
310                         rsPayment!Status = "Good"
312                         rsPayment.Update
                        End If

314                     rsCustomer.Update
            
316                     If rsTrail.State = 1 Then rsTrail.Close
318                     rsTrail.Open "Select * from tblTrail "
                    
320                     With rsTrail
322                         .AddNew
324                         !UserName = lblUser.Caption
326                         !userlevel = lblUserlevel.Caption
328                         !Activity = "Reverse payment"
330                         !Time = lblTime.Caption
332                         !Date = txtTodayDate.Text
334                         .Update
                        End With
            
336                     MsgBox "Record Successfully Reversed", vbInformation, _
                                "Webplus Lending Corporation"
338                     Call auditPayment(LoanID)
                        ' Unload Me
340                     frm_PaymentReverse.lblUser.Caption = MDIForm1.lblUserName.Caption
342                     Me.Show
                    
                    End If

344                 Call auditPayment(LoanID)
346                 txtSearch.Enabled = True
                End If
            End If
        End If

348     Call setgrid2

        '<EhFooter>

        TxtLog "Exited btnReverse_Click"

        Exit Sub

btnReverse_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_PaymentReverse.btnReverse_Click", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnReverse_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo btnReverse_KeyPress_Err

        TxtLog "Entered btnReverse_KeyPress"

        '</EhHeader>

100     KeyAscii = 0

        '<EhFooter>

        TxtLog "Exited btnReverse_KeyPress"

        Exit Sub

btnReverse_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_PaymentReverse.btnReverse_KeyPress", Erl

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
        ErrReport Err.Description, _
                "LendingClient.frm_PaymentReverse.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call payment
104     Call CollData
106     Call Loan
108     Call Loan1
110     Call Customer
        
        'to create new OR for reverse
112     Call autoNumber

114     If rsPayment.State = 1 Then rsPayment.Close
116     rsPayment.Open "Select TOP 25 * from tblPayment Order By ID desc"
118     Set DataGrid1.DataSource = rsPayment
120     DataGrid1.Width = Me.Width
122     DataGrid1.Columns(0).Width = 0
124     DataGrid1.Columns(1).Width = 1100
126     DataGrid1.Columns(3).Width = 900
128     DataGrid1.Columns(4).Width = 900
130     DataGrid1.Columns(6).Width = 1200
132     DataGrid1.Columns(7).Width = 1000
134     DataGrid1.Columns(8).Width = 1000
136     DataGrid1.Columns(9).Width = 1100
138     DataGrid1.Columns(10).Width = 1100
140     DataGrid1.Columns(12).Width = 1100
142     DataGrid1.Columns(15).Width = 1100

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_PaymentReverse.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    txtTodayDate.Text = Date
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

110         If rsPayment.State = 1 Then rsPayment.Close
112         rsPayment.Open "Select * from tblPayment where Code = '" & txtCode.Text & "'"

114         If rsPayment.RecordCount = 0 Then
116             MsgBox "No Record Found", vbInformation, "Webplus Lending Corporation"
118             txtCode.Text = ""
120             txtCode.SetFocus
            Else

122             With rsPayment
124                 txtCode.Text = !code
126                 txtCollector.Text = !Collector
128                 txtCustomer.Text = !Customer
130                 txtPrincipal.Text = !principal
132                 txtAmortization.Text = !Amortization
134                 txtDateRelease.Text = !DateRelease
136                 txtMaturity.Text = !Maturity
138                 txtTotaltAmortization.Text = !TotalBalance
140                 txtOutBalance.Text = !TotalBalance
142                 txtPaymentmade.Text = !paymentsMade
144                 txtAmountPaid.Text = !paymentsMade
146                 txtNBalance.Text = !NewBalance
148                 txtTotaltBalance.Text = !NewBalance
                End With

            End If

        Else
150         KeyAscii = 0
    
        End If

        '<EhFooter>

        TxtLog "Exited txtCode_KeyPress"

        Exit Sub

txtCode_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_PaymentReverse.txtCode_KeyPress", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtOR_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtOR_KeyPress_Err

        TxtLog "Entered txtOR_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0123456789"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
        
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf KeyAscii = 13 Then
            
            ' Call lenght

            ' lblOR.Caption = txtR.Text + txtOR.Text
110         lblOR.Caption = txtOR.Text
            
112         If rsPayment.State = 1 Then rsPayment.Close
114         rsPayment.Open "Select * from tblPayment where ORnumber = " & lblOR.Caption _
                    & " and (Status = 'Good' or Status = 'Full Paid')"
            
116         If rsPayment.RecordCount <> 0 Then
                
118             With rsPayment
120                 txtLoanID.Text = !LoanID
122                 txtCode.Text = !code
124                 txtCollector.Text = !Collector
126                 txtFCollector.Text = !CollectorFname
128                 txtCustomer.Text = !Customer
130                 txtPrincipal.Text = !principal
132                 txtAmortization.Text = !Amortization
134                 txtDateRelease.Text = !DateRelease
136                 txtMaturity.Text = !Maturity
138                 txtPaymentmade.Text = !paymentsMade
140                 lblPaymentMade.Caption = !paymentsMade
142                 txtAmountPaid.Text = !paymentsMade
144                 lblTotalBalance.Caption = !TotalBalance
146                 txtNBalance.Text = !NewBalance
148                 txtTotaltBalance.Text = !NewBalance
150                 lblCollectorCode.Caption = !CollectorCode
152                 lblTotalPayment.Caption = !TotalPayment
154                 lblPaOR.Caption = lblOR.Caption
                End With
            
156             If rsLoan.State = 1 Then rsLoan.Close
158             rsLoan.Open "Select * from tblLoan where LoanID = " & txtLoanID.Text & " "

160             If rsLoan.RecordCount <> 0 Then
162                 txtTotaltAmortization.Text = rsLoan!TotalAmortization
164                 txtOutBalance.Text = rsLoan!TotalAmortization
166                 lblTotalPayment.Caption = rsLoan!TotalPayment
                End If

168             txtTotaltBalance.Text = Val(txtTotaltAmortization.Text) + Val( _
                        txtPaymentmade.Text)
170             txtNBalance.Text = txtTotaltBalance.Text
172             txtPaymentmade.Text = lblminus.Caption + txtPaymentmade.Text
            Else
174             MsgBox "Payment Already  Reversed or No Record Found", vbInformation, _
                        "Webplus Lending Corporation"
            End If

        Else
176         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtOR_KeyPress"

        Exit Sub

txtOR_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_PaymentReverse.txtOR_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsPayment.State = 1 Then rsPayment.Close
102     rsPayment.Open "Select * from tblPayment where Customer like '" & _
                txtSearch.Text & "%' or LoanID like '" & txtSearch.Text & _
                "%' or ORnumber like '" & txtSearch.Text & "%' or DateEncoded like '" & _
                txtSearch.Text & "%' or Code like '" & txtSearch.Text & _
                "%' Order By ORnumber desc"
        
104     Set DataGrid1.DataSource = rsPayment
106     DataGrid1.Width = Me.Width
108     DataGrid1.Columns(0).Width = 0
110     DataGrid1.Columns(1).Width = 1100
112     DataGrid1.Columns(3).Width = 900
114     DataGrid1.Columns(4).Width = 900
116     DataGrid1.Columns(6).Width = 1200
118     DataGrid1.Columns(7).Width = 1000
120     DataGrid1.Columns(8).Width = 1000
122     DataGrid1.Columns(9).Width = 1100
124     DataGrid1.Columns(10).Width = 1100
126     DataGrid1.Columns(12).Width = 1100
128     DataGrid1.Columns(15).Width = 1100

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_PaymentReverse.txtsearch_Change", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

