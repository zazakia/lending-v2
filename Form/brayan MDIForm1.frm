VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H0080FFFF&
   Caption         =   "Lending Business Management System by www.teamwebplus.com"
   ClientHeight    =   8130
   ClientLeft      =   -20820
   ClientTop       =   1080
   ClientWidth     =   14820
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   16440
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":1064
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":2558
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":2DF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":3A15
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":6134
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":7FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":9A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":AD5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":B39A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":D659
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":ECC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "brayan MDIForm1.frx":481D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   2190
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   3863
      ButtonWidth     =   2910
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Customers (Ctrl+C)"
            Key             =   "sss"
            Object.Tag             =   "sss"
            ImageIndex      =   12
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Customers"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Loans (Ctrl+L)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Payments  (Ctrl+P)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DCR  (Ctrl+D)"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CS  (Ctrl+A)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reverse Payments"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reverse Loan"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Collector"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "VCL"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "VCP"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Dashboard"
            Description     =   "KPI"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   14820
      TabIndex        =   22
      Top             =   2190
      Width           =   14820
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   14820
      TabIndex        =   26
      Top             =   2190
      Width           =   14820
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H0080FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5940
      Left            =   0
      ScaleHeight     =   5880
      ScaleWidth      =   2.45685e5
      TabIndex        =   0
      Top             =   2190
      Width           =   2.45745e5
      Begin VB.CommandButton cmdRefreshDashboard 
         Caption         =   "Refresh Dashboard Data"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   11280
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   261423105
         CurrentDate     =   41925
      End
      Begin VB.CommandButton cmdExportReport 
         Caption         =   "ExportReport"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   6240
         TabIndex        =   12
         Top             =   7080
         Visible         =   0   'False
         Width           =   6615
      End
      Begin MSComCtl2.DTPicker dtLogdate 
         Height          =   375
         Left            =   19080
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Format          =   261423105
         CurrentDate     =   41869
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   17280
         Top             =   720
      End
      Begin VB.Shape Shape2 
         Height          =   2775
         Left            =   7920
         Top             =   2640
         Width           =   9615
      End
      Begin VB.Shape Shape1 
         Height          =   4455
         Left            =   120
         Top             =   1440
         Width           =   18135
      End
      Begin VB.Label lblRateOf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate of Return = Total Payments / Total Principal x 100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8400
         TabIndex        =   25
         Top             =   2880
         Width           =   8715
      End
      Begin VB.Label lbldfgfd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   10560
         TabIndex        =   24
         Top             =   3480
         Width           =   2280
      End
      Begin VB.Label lblRateOfReturn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblRateOfReturn"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1665
         Left            =   8760
         TabIndex        =   23
         Top             =   3480
         Width           =   11040
      End
      Begin VB.Label lblNumberOfLoans 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NumberOfLoans"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   480
         TabIndex        =   21
         Top             =   4080
         Width           =   3255
      End
      Begin VB.Label lblNumberofPayments 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NumberofPayments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   480
         TabIndex        =   20
         Top             =   5160
         Width           =   3900
      End
      Begin VB.Label lblTotalPayments 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sum of Payments Made for All Clients"
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
         Left            =   240
         TabIndex        =   19
         Top             =   4680
         Width           =   7560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sum of All Loans Based on Principal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   240
         TabIndex        =   18
         Top             =   3720
         Width           =   7350
      End
      Begin VB.Label lblNumberOf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Number of Clients (All Clients)"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   6315
      End
      Begin VB.Label lblNumberOfClients 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Clients"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   540
         Left            =   480
         TabIndex        =   15
         Top             =   3000
         Width           =   3555
      End
      Begin VB.Label lblVersion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   10680
         TabIndex        =   14
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9/17/2014"
         Height          =   195
         Left            =   15000
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblLabel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome      :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   10
         Top             =   6240
         Width           =   1965
      End
      Begin VB.Label lblUserlevel 
         Caption         =   "userlevel"
         Height          =   495
         Left            =   17760
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblusername 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   690
         Left            =   2640
         TabIndex        =   6
         Top             =   6120
         Width           =   1890
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DAILY COLLECTION AND LOAN REPORT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   435
         Left            =   5805
         TabIndex        =   5
         Top             =   600
         Width           =   7425
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MELANN LENDING INVESTOR CORP."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   5325
         TabIndex        =   4
         Top             =   0
         Width           =   8535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today is : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label lbltime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   645
      End
      Begin VB.Label lbldate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   570
      End
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu CP 
         Caption         =   "ChangePassword"
      End
      Begin VB.Menu Logout 
         Caption         =   "Logout"
      End
   End
   Begin VB.Menu Ed 
      Caption         =   "&Edit"
      Begin VB.Menu CU 
         Caption         =   "Customer"
         Shortcut        =   ^C
      End
      Begin VB.Menu co 
         Caption         =   "Collector"
      End
      Begin VB.Menu users 
         Caption         =   "Users"
      End
      Begin VB.Menu Branch 
         Caption         =   "Branch"
      End
      Begin VB.Menu Cashonhand 
         Caption         =   "Cash On Hand"
      End
      Begin VB.Menu delivery 
         Caption         =   "Charges"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu colOvers 
         Caption         =   "Collector Overs"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu Ch 
         Caption         =   "Chart of Accounts"
      End
   End
   Begin VB.Menu tr 
      Caption         =   "&Transaction"
      Begin VB.Menu L 
         Caption         =   "&Loan"
         Shortcut        =   ^L
      End
      Begin VB.Menu payment 
         Caption         =   "Payment"
         Shortcut        =   ^P
      End
      Begin VB.Menu Rev 
         Caption         =   "Reverse Loan"
      End
      Begin VB.Menu ReversePayment 
         Caption         =   "Reverse Payment"
      End
      Begin VB.Menu Break 
         Caption         =   "Breakdown"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu llllll 
         Caption         =   "-"
      End
      Begin VB.Menu depositsand 
         Caption         =   "Deposits and Other"
         Shortcut        =   ^B
      End
      Begin VB.Menu expe 
         Caption         =   "Expense"
      End
   End
   Begin VB.Menu Inq 
      Caption         =   "Inquiry"
      Begin VB.Menu viewwedin 
         Caption         =   "View Customer Loan"
      End
      Begin VB.Menu viewlang 
         Caption         =   "View Customer Payment"
      End
      Begin VB.Menu tra 
         Caption         =   "Trail"
      End
      Begin VB.Menu ytyty 
         Caption         =   "-"
      End
      Begin VB.Menu ViewC 
         Caption         =   "View Collections"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu rep 
      Caption         =   "Reports"
      Begin VB.Menu LoanList 
         Caption         =   "LoanList"
      End
      Begin VB.Menu dcr 
         Caption         =   "Daily Cash Report"
         Shortcut        =   ^D
      End
      Begin VB.Menu coll 
         Caption         =   "Collection Reports"
         Shortcut        =   ^A
      End
      Begin VB.Menu CollectorList 
         Caption         =   "Collector List"
      End
      Begin VB.Menu CPR 
         Caption         =   "Statement of Account"
      End
      Begin VB.Menu pay_en_menu 
         Caption         =   "Payments Encoded"
      End
      Begin VB.Menu pay_rev_menu 
         Caption         =   "Payments Reversed"
      End
      Begin VB.Menu inact_cos 
         Caption         =   "Inactive Costumer Report"
      End
      Begin VB.Menu mn_full_paid 
         Caption         =   "Fully Paid as of Date Reports"
      End
      Begin VB.Menu mn_monthly_collections 
         Caption         =   "Monthly Collections"
      End
      Begin VB.Menu mn_monthlyReleased 
         Caption         =   "Monthly Released"
      End
      Begin VB.Menu Auditor 
         Caption         =   "Audit Loan Checker"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MDIForm1
'    Project    : LendingClient
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Private Sub CloseChildMDI()

        '<EhHeader>
        On Error GoTo CloseChildMDI_Err

        TxtLog "Entered CloseChildMDI"

        '</EhHeader>

100     Unload frm_payment
    
102     Unload frm_Loan
    
        '<EhFooter>

        TxtLog "Exited CloseChildMDI"

        Exit Sub

CloseChildMDI_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.CloseChildMDI", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       ComputeNumberOfClients
' Description:       Dash board Totals and Computation for stats
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       1/1/2018-8:44:13 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub ComputeNumberOfClients()

        '<EhHeader>
        On Error GoTo ComputeNumberOfClients_Err

        TxtLog "Entered ComputeNumberOfClients"

        '</EhHeader>

100     Call connect

        'Calls recordset for Customer table
102     Set rsCustomer = Nothing
104     Set rsCustomer = New adodb.Recordset

106     With rsCustomer
108         .CursorType = adOpenDynamic
110         .LockType = adLockOptimistic
112         .ActiveConnection = conn
114         .Source = "Select * from tblCustomer "
116         .CursorLocation = adUseClient
118         .Open
        End With

120     lblNumberOfClients = rsCustomer.RecordCount

122     Set rsLoan = Nothing
124     Set rsLoan = New adodb.Recordset

126     With rsLoan
128         .CursorType = adOpenDynamic
130         .LockType = adLockOptimistic
132         .ActiveConnection = conn
134         .Source = _
                    "SELECT DISTINCTROW Sum(tblLoan.Principal) AS SumOfPaymentsMade FROM tblLoan;"
136         .CursorLocation = adUseClient
138         .Open
        End With

140     lblNumberOfLoans = FormatCurrency(rsLoan.Fields(0))
        
142     Set rsPayment = Nothing
        
144     Set rsPayment = New adodb.Recordset

146     With rsPayment
148         .CursorType = adOpenDynamic
150         .LockType = adLockOptimistic
152         .ActiveConnection = conn
154         .Source = _
                    "SELECT DISTINCTROW Sum(tblPayment.PaymentsMade) AS SumOfPaymentsMade FROM tblPayment;"
156         .CursorLocation = adUseClient
158         .Open
        End With

160     lblNumberofPayments = FormatCurrency(rsPayment.Fields(0))

162     lblRateOfReturn = Round(Val(Format(lblNumberofPayments)) / Val(Format( _
                lblNumberOfLoans)) * 100, 0)

        '<EhFooter>

        TxtLog "Exited ComputeNumberOfClients"

        Exit Sub

ComputeNumberOfClients_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.ComputeNumberOfClients", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Auditor_Click()

        '<EhHeader>
        On Error GoTo Auditor_Click_Err

        TxtLog "Entered Auditor_Click"

        '</EhHeader>

        Dim LoanID As Single

100     LoanID = InputBox("Ente Loan ID")
    
102     auditPayment (LoanID)
    
        '<EhFooter>

        TxtLog "Exited Auditor_Click"

        Exit Sub

Auditor_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.Auditor_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Branch_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Branch_Click()

        '<EhHeader>
        On Error GoTo Branch_Click_Err

        TxtLog "Entered Branch_Click"

        '</EhHeader>

100     Load frm_Branch
102     frm_Branch.Show

        '<EhFooter>

        TxtLog "Exited Branch_Click"

        Exit Sub

Branch_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.Branch_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Break_Click
' Description:       money breakdown feature
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Break_Click()

    '<EhHeader>
    On Error GoTo Break_Click_Err

    TxtLog "Entered Break_Click"

    '</EhHeader>

    '<EhFooter>

    TxtLog "Exited Break_Click"

    Exit Sub

Break_Click_Err:
    ErrReport Err.Description, "LendingClient.MDIForm1.Break_Click", Erl

    Resume Next

    '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Cashonhand_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Cashonhand_Click()

        '<EhHeader>
        On Error GoTo Cashonhand_Click_Err

        TxtLog "Entered Cashonhand_Click"

        '</EhHeader>

100     frm_CashOnHand.lblUser.Caption = MDIForm1.lblUserName.Caption
102     Load frm_CashOnHand
104     frm_CashOnHand.Show

        '<EhFooter>

        TxtLog "Exited Cashonhand_Click"

        Exit Sub

Cashonhand_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.Cashonhand_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Ch_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Ch_Click()

        '<EhHeader>
        On Error GoTo Ch_Click_Err

        TxtLog "Entered Ch_Click"

        '</EhHeader>

100     Load frm_Chart
102     frm_Chart.Show

        '<EhFooter>

        TxtLog "Exited Ch_Click"

        Exit Sub

Ch_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.Ch_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       cmdExportReport_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub cmdExportReport_Click()

        '<EhHeader>
        On Error GoTo cmdExportReport_Click_Err

        TxtLog "Entered cmdExportReport_Click"

        '</EhHeader>
        
        Dim CRApp            As New CRAXDDRT.Application

        Dim Report           As New CRAXDDRT.Report

        Dim crxExportOptions As CRAXDRT.ExportOptions

100     Set CRApp = New CRAXDRT.Application
    
        'Report.RecordSelectionFormula = "({tblLoan.DateRelease} in Date(#" & dtpFrom & "#) to Date(#" & dtpTo & "#)) and ({tblLoan.Status} = 'Good')"
        'Report.RecordSelectionFormula = "{tblLoan.Status} = 'Good'"
        
        'test stuff start
       
        'convert date to string start

        Dim testDate        As Date

        Dim testMonthString As String

        Dim testDayInt      As Integer, testMonthInt As Integer, testYearInt As Integer
    
102     testDate = DTPicker1
        
        'next time this will be the date from the date picker.
        'because the date is supposed to be chosen by the user and not constant.
        'then a path based on the directory will be created based on the chosen date.

104     testDayInt = Day(testDate)
106     MsgBox (testDayInt)
108     testMonthInt = Month(testDate)
110     MsgBox (testMonthInt)
112     testMonthString = MonthName(testMonthInt)
114     MsgBox (testMonthString)
116     testYearInt = Year(testDate)
118     MsgBox (testYearInt)
    
        'directory checker and creator code start
    
        'If Dir$("D:\Dropbox\tblBranch1"), vbDirectory) = vbNullString Then --- the branch directory should be based on the database and if does not exist it will be created
        'MkDir ("D:\Dropbox\tblBranch1")
        'MsgBox ("The directory doesn't exist. The directory was created instead.")
        'end if
    
120     If rsBranch.State = 1 Then rsBranch.Close
122     rsBranch.Open "Select * from tblBranch"
    
        Dim exportDir As String

124     exportDir = "D:\DropBox\"
    
126     If Dir$(exportDir, vbDirectory) = vbNullString Then 'turn this into a function
128         MkDir (exportDir)
130         MsgBox ("The direcotry doesn't exist. The directory was created instead.")
        Else
132         MsgBox ("Directory already exists.")
        End If
    
134     exportDir = exportDir & rsBranch!BranchName
    
136     If Dir$(exportDir, vbDirectory) = vbNullString Then 'turn this into a function
138         MkDir (exportDir)
140         MsgBox ("The direcotry doesn't exist. The directory was created instead.")
        Else
142         MsgBox ("Directory already exists.")
        End If
    
144     exportDir = exportDir & "\reports"
    
146     If Dir$(exportDir, vbDirectory) = vbNullString Then 'turn this into a function
148         MkDir (exportDir)
150         MsgBox ("The direcotry doesn't exist. The directory was created instead.")
        Else
152         MsgBox ("Directory already exists.")
        End If
    
154     exportDir = exportDir & "\" & CStr(testYearInt)

156     If Dir$(exportDir, vbDirectory) = vbNullString Then
158         MkDir (exportDir)
160         MsgBox ("The direcotry doesn't exist. The directory was created instead.")
        Else
162         MsgBox ("Directory already exists.")
        End If
    
164     exportDir = exportDir & "\" & testMonthString

166     If Dir$(exportDir, vbDirectory) = vbNullString Then
168         MkDir (exportDir)
170         MsgBox ("The direcotry doesn't exist. The directory was created instead.")
        Else
172         MsgBox ("Directory already exists.")
        End If
    
174     exportDir = exportDir & "\" & CStr(testDayInt)

176     If Dir$(exportDir, vbDirectory) = vbNullString Then
178         MkDir (exportDir)
180         MsgBox ("The direcotry doesn't exist. The directory was created instead.")
        Else
182         MsgBox ("Directory already exists.")
        End If
    
        'directory checker and creator code end
    
184     Set Report = CRApp.OpenReport(App.Path & "\report\DSR LOANS AS MAIN.rpt ")
186     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
188     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
        'CRViewer.ReportSource = Report
190     Report.RecordSelectionFormula = "({tblLoan.DateRelease} in Date(#" & DTPicker1 _
                & "#) to Date(#" & DTPicker1 & _
                "#)) and ({tblLoan.Status} = 'Good' or {tblLoan.Status} = 'Full Paid')"

192     Set crxExportOptions = Report.ExportOptions
194     crxExportOptions.DestinationType = crEDTDiskFile
196     crxExportOptions.DiskFileName = exportDir & "\dcr.pdf"
198     crxExportOptions.FormatType = crEFTPortableDocFormat
200     crxExportOptions.PDFFirstPageNumber = 1
202     crxExportOptions.PDFLastPageNumber = 1
204     crxExportOptions.PDFExportAllPages = True
206     Report.Export False
       
208     Set Report = CRApp.OpenReport(App.Path & "\report\collections.rpt")
210     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
212     Report.RecordSelectionFormula = "{tblLoan.DateRelease}  < #" & DTPicker1 & _
                "# and {tblLoan.Status} = 'Good' and {tblCustomer.Balance} <> 0"
        'Report.RecordSelectionFormula = "{tblLoan.DateRelease}  < #" & dtp_cs & "# and {tblLoan.Status} = 'Good' and {tblCustomer.Balance} <> 0"
    
214     Set crxExportOptions = Report.ExportOptions
216     crxExportOptions.DestinationType = crEDTDiskFile
218     crxExportOptions.DiskFileName = exportDir & "\collection sheet.pdf"
220     crxExportOptions.FormatType = crEFTPortableDocFormat
222     crxExportOptions.PDFFirstPageNumber = 1
224     crxExportOptions.PDFLastPageNumber = 1
226     crxExportOptions.PDFExportAllPages = True
228     Report.Export False
        
        'batch 3 end

        '<EhFooter>

        TxtLog "Exited cmdExportReport_Click"

        Exit Sub

cmdExportReport_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.cmdExportReport_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdRefreshDashboard_Click()

        '<EhHeader>
        On Error GoTo cmdRefreshDashboard_Click_Err

        TxtLog "Entered cmdRefreshDashboard_Click"

        '</EhHeader>

100     Call ComputeNumberOfClients

        '<EhFooter>

        TxtLog "Exited cmdRefreshDashboard_Click"

        Exit Sub

cmdRefreshDashboard_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.cmdRefreshDashboard_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Co_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Co_Click()

        '<EhHeader>
        On Error GoTo Co_Click_Err

        TxtLog "Entered Co_Click"

        '</EhHeader>

        'frm_Collectors.lblUserlevel.Caption = MDIForm1.lblUserlevel.Caption
        'frm_Collectors.lblUserr.Caption = MDIForm1.lblusername.Caption
100     frm_Collectors.lblUserName.Caption = " " & MDIForm1.lblUserName.Caption
102     Load frm_Collectors
104     frm_Collectors.Show

        '<EhFooter>

        TxtLog "Exited Co_Click"

        Exit Sub

Co_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.Co_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Coll_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Coll_Click()

        '<EhHeader>
        On Error GoTo Coll_Click_Err

        TxtLog "Entered Coll_Click"

        '</EhHeader>

100     Unload rep_CollectionSheet
102     Load rep_CollectionSheet
104     rep_CollectionSheet.Show

        '<EhFooter>

        TxtLog "Exited Coll_Click"

        Exit Sub

Coll_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.Coll_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       CollectorList_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CollectorList_Click()

        '<EhHeader>
        On Error GoTo CollectorList_Click_Err

        TxtLog "Entered CollectorList_Click"

        '</EhHeader>

100     Load rep_CollectorList
102     rep_CollectorList.Show

        '<EhFooter>

        TxtLog "Exited CollectorList_Click"

        Exit Sub

CollectorList_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.CollectorList_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       colOvers_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub colOvers_Click()

        '<EhHeader>
        On Error GoTo colOvers_Click_Err

        TxtLog "Entered colOvers_Click"

        '</EhHeader>

100     frm_CollectorOvers.lblUserlevel.Caption = MDIForm1.lblUserlevel.Caption
102     frm_CollectorOvers.lblUserr.Caption = MDIForm1.lblUserName.Caption
104     frm_CollectorOvers.txtUser.Text = MDIForm1.lblUserName.Caption
106     Load frm_CollectorOvers
108     frm_CollectorOvers.Show

        '<EhFooter>

        TxtLog "Exited colOvers_Click"

        Exit Sub

colOvers_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.colOvers_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       CP_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CP_Click()

        '<EhHeader>
        On Error GoTo CP_Click_Err

        TxtLog "Entered CP_Click"

        '</EhHeader>

100     Load frm_ChangePassword
102     frm_ChangePassword.Show

        '<EhFooter>

        TxtLog "Exited CP_Click"

        Exit Sub

CP_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.CP_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       CPR_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CPR_Click()

        '<EhHeader>
        On Error GoTo CPR_Click_Err

        TxtLog "Entered CPR_Click"

        '</EhHeader>

100     Unload rep_cprForm
102     Load rep_cprForm
104     rep_cprForm.Show

        '<EhFooter>

        TxtLog "Exited CPR_Click"

        Exit Sub

CPR_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.CPR_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       CU_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub CU_Click()

        '<EhHeader>
        On Error GoTo CU_Click_Err

        TxtLog "Entered CU_Click"

        '</EhHeader>

100     Picture1.Visible = False
102     frm_Customer.lblUser.Caption = MDIForm1.lblUserName.Caption
104     frm_Customer.txtUser.Text = MDIForm1.lblUserName.Caption
106     frm_Customer.lblUserlevel.Caption = rsUser!userlevel
108     Load frm_Customer
110     frm_Customer.Show
112     frm_Customer.SetFocus

        '<EhFooter>

        TxtLog "Exited CU_Click"

        Exit Sub

CU_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.CU_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub dcr_Click()

        '<EhHeader>
        On Error GoTo dcr_Click_Err

        TxtLog "Entered dcr_Click"

        '</EhHeader>
        
100     Unload rep_DailySalesReport
102     Load rep_DailySalesReport
104     rep_DailySalesReport.Show

        '<EhFooter>

        TxtLog "Exited dcr_Click"

        Exit Sub

dcr_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.dcr_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       delivery_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub delivery_Click()

        '<EhHeader>
        On Error GoTo delivery_Click_Err

        TxtLog "Entered delivery_Click"

        '</EhHeader>

100     Load frm_Delivery
102     frm_Delivery.Show

        '<EhFooter>

        TxtLog "Exited delivery_Click"

        Exit Sub

delivery_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.delivery_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       depositsand_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub depositsand_Click()

        '<EhHeader>
        On Error GoTo depositsand_Click_Err

        TxtLog "Entered depositsand_Click"

        '</EhHeader>

100     frm_Deposit.lblUser.Caption = MDIForm1.lblUserName.Caption
102     frm_Deposit.txtUser.Text = MDIForm1.lblUserName.Caption
104     frm_Deposit.lblUserlevel.Caption = MDIForm1.lblUserlevel.Caption
106     Load frm_Deposit
108     frm_Deposit.Show

        '<EhFooter>

        TxtLog "Exited depositsand_Click"

        Exit Sub

depositsand_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.depositsand_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       expe_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub expe_Click()

        '<EhHeader>
        On Error GoTo expe_Click_Err

        TxtLog "Entered expe_Click"

        '</EhHeader>

        ' frm_Expense.lblUser.Caption = rsUser!UserName
100     frm_Expenses.txtUser.Text = MDIForm1.lblUserName.Caption
102     Load frm_Expenses
104     frm_Expenses.Show

        '<EhFooter>

        TxtLog "Exited expe_Click"

        Exit Sub

expe_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.expe_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       inact_cos_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub inact_cos_Click()

        '<EhHeader>
        On Error GoTo inact_cos_Click_Err

        TxtLog "Entered inact_cos_Click"

        '</EhHeader>

100     Unload rep_inactivecleints
102     Load rep_inactivecleints
104     rep_inactivecleints.Show

        '<EhFooter>

        TxtLog "Exited inact_cos_Click"

        Exit Sub

inact_cos_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.inact_cos_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       L_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub L_Click()

        '<EhHeader>
        On Error GoTo L_Click_Err

        TxtLog "Entered L_Click"

        '</EhHeader>

100     Picture1.Visible = False
102     frm_Loan.lblUser.Caption = MDIForm1.lblUserName.Caption
104     frm_Loan.lblUserlevel.Caption = rsUser!userlevel
        'Load frm_Loan
        'frm_Loan.Show
106     frm_Loan.SetFocus

        '<EhFooter>

        TxtLog "Exited L_Click"

        Exit Sub

L_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.L_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       LoanList_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub LoanList_Click()

        '<EhHeader>
        On Error GoTo LoanList_Click_Err

        TxtLog "Entered LoanList_Click"

        '</EhHeader>

100     Load frm_CustomerLoan
102     frm_CustomerLoan.Show

        '<EhFooter>

        TxtLog "Exited LoanList_Click"

        Exit Sub

LoanList_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.LoanList_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Logout_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Logout_Click()

        '<EhHeader>
        On Error GoTo Logout_Click_Err

        TxtLog "Entered Logout_Click"

        '</EhHeader>

100     If MsgBox("Do you want to Logout?", vbQuestion + vbYesNo, _
                "Webplus Lending Corporation") = vbYes Then

102         If rsLogin.State = 1 Then rsLogin.Close
104         rsLogin.Open "select * from tblUser where Username = '" & _
                    lblUserName.Caption & "'"

106         If rsLogin.RecordCount <> 0 Then

108             With rsLogin
110                 !Status = "Log-out"
112                 .Update
                End With

            Else
114             MsgBox "    ", vbInformation, "Webplus Lending Corporation"
            End If

116         rsTrail.AddNew
118         rsTrail!userlevel = lblUserlevel.Caption
120         rsTrail!UserName = Trim$(lblUserName.Caption)
122         rsTrail!Activity = "Logged Out"
124         rsTrail!Time = lblTime.Caption
126         rsTrail!Date = dtLogdate.Value
128         rsTrail!Status = "Log-out"
130         rsTrail.Update
    
132         MDIForm1.lblUserName = ""
134         frm_Customer.lblUser.Caption = ""
136         Unload Me
138         frmLogin.Show
        
        End If

        '<EhFooter>

        TxtLog "Exited Logout_Click"

        Exit Sub

Logout_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.Logout_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       MDIForm_Load
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub MDIForm_Load()

        '<EhHeader>
        On Error GoTo MDIForm_Load_Err

        TxtLog "Entered MDIForm_Load"

        '</EhHeader>

        'Call connect
100     Call Login
102     Call connect
104     lblVersion = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
        
106     Set rsBranch = Nothing
108     Set rsBranch = New adodb.Recordset

110     With rsBranch
112         .CursorType = adOpenDynamic
114         .LockType = adLockOptimistic
116         .ActiveConnection = conn
118         .Source = "Select * from tblBranch"
120         .CursorLocation = adUseClient
122         .Open
        End With
        
124     Call ComputeNumberOfClients
        '106     MDIForm1.WindowState = 2

        '<EhFooter>

        TxtLog "Exited MDIForm_Load"

        Exit Sub

MDIForm_Load_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.MDIForm_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       MDIForm_Unload
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       Cancel (Integer)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub MDIForm_Unload(Cancel As Integer)

        '<EhHeader>
        On Error GoTo MDIForm_Unload_Err

        TxtLog "Entered MDIForm_Unload"

        '</EhHeader>

100     UnloadAllForms
        'Unload Me

        '<EhFooter>

        TxtLog "Exited MDIForm_Unload"

        Exit Sub

MDIForm_Unload_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.MDIForm_Unload", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       mn_full_paid_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub mn_full_paid_Click()

        '<EhHeader>
        On Error GoTo mn_full_paid_Click_Err

        TxtLog "Entered mn_full_paid_Click"

        '</EhHeader>

100     Unload rep_full_paid
102     Load rep_full_paid
104     rep_full_paid.Show

        '<EhFooter>

        TxtLog "Exited mn_full_paid_Click"

        Exit Sub

mn_full_paid_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.mn_full_paid_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       mn_monthly_collections_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub mn_monthly_collections_Click()

        '<EhHeader>
        On Error GoTo mn_monthly_collections_Click_Err

        TxtLog "Entered mn_monthly_collections_Click"

        '</EhHeader>

100     Unload rep_month_collection
102     Load rep_month_collection
104     rep_month_collection.Show

        '<EhFooter>

        TxtLog "Exited mn_monthly_collections_Click"

        Exit Sub

mn_monthly_collections_Click_Err:
        ErrReport Err.Description, _
                "LendingClient.MDIForm1.mn_monthly_collections_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       mn_monthlyReleased_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub mn_monthlyReleased_Click()

        '<EhHeader>
        On Error GoTo mn_monthlyReleased_Click_Err

        TxtLog "Entered mn_monthlyReleased_Click"

        '</EhHeader>

100     Unload rep_month_released
102     Load rep_month_released
104     rep_month_released.Show

        '<EhFooter>

        TxtLog "Exited mn_monthlyReleased_Click"

        Exit Sub

mn_monthlyReleased_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.mn_monthlyReleased_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       pay_en_menu_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub pay_en_menu_Click()

        '<EhHeader>
        On Error GoTo pay_en_menu_Click_Err

        TxtLog "Entered pay_en_menu_Click"

        '</EhHeader>

100     Unload rep_PaymentsEncoded
102     Load rep_PaymentsEncoded
104     rep_PaymentsEncoded.Show

        '<EhFooter>

        TxtLog "Exited pay_en_menu_Click"

        Exit Sub

pay_en_menu_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.pay_en_menu_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       pay_rev_menu_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub pay_rev_menu_Click()

        '<EhHeader>
        On Error GoTo pay_rev_menu_Click_Err

        TxtLog "Entered pay_rev_menu_Click"

        '</EhHeader>

100     Unload rep_PaymentsReversed
102     Load rep_PaymentsReversed
104     rep_PaymentsReversed.Show

        '<EhFooter>

        TxtLog "Exited pay_rev_menu_Click"

        Exit Sub

pay_rev_menu_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.pay_rev_menu_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       payment_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub payment_Click()

        '<EhHeader>
        On Error GoTo payment_Click_Err

        TxtLog "Entered payment_Click"

        '</EhHeader>

100     Picture1.Visible = False
        'frm_payment.btnAddpayment.SetFocus
102     frm_payment.lblUser.Caption = MDIForm1.lblUserName.Caption
104     frm_payment.lblUserlevel.Caption = MDIForm1.lblUserlevel.Caption
        ' Load frm_payment
106     frm_payment.Show
108     frm_payment.SetFocus

        '<EhFooter>

        TxtLog "Exited payment_Click"

        Exit Sub

payment_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.payment_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Rev_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Rev_Click()

        '<EhHeader>
        On Error GoTo Rev_Click_Err

        TxtLog "Entered Rev_Click"

        '</EhHeader>

100     frm_LoanReverse.lblUser.Caption = MDIForm1.lblUserName.Caption
102     frm_LoanReverse.lblUserlevel.Caption = MDIForm1.lblUserlevel.Caption
104     Load frm_LoanReverse
106     frm_LoanReverse.Show

        '<EhFooter>

        TxtLog "Exited Rev_Click"

        Exit Sub

Rev_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.Rev_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       ReversePayment_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub ReversePayment_Click()

        '<EhHeader>
        On Error GoTo ReversePayment_Click_Err

        TxtLog "Entered ReversePayment_Click"

        '</EhHeader>

100     frm_PaymentReverse.lblUser.Caption = MDIForm1.lblUserName.Caption
102     Load frm_PaymentReverse
104     frm_PaymentReverse.Show

        '<EhFooter>

        TxtLog "Exited ReversePayment_Click"

        Exit Sub

ReversePayment_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.ReversePayment_Click", Erl

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

    lbldate.Caption = FormatDateTime(Date, vbLongDate)
    
    lblTime.Caption = Time

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Toolbar1_ButtonClick
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :       Button (MSComctlLib.Button)
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

        '<EhHeader>
        On Error GoTo Toolbar1_ButtonClick_Err

        TxtLog "Entered Toolbar1_ButtonClick"

        '</EhHeader>

100     Select Case Button.Index
    
            Case 1 'CUSTOMER
                
102             CU_Click

104         Case 2 'LOAN
                
106             L_Click

108         Case 3 'PAYMENTS
                
110             payment_Click

112         Case 4
                
114             dcr_Click

116         Case 5
118             Picture1.Visible = False
120             Coll_Click

122         Case 6
124             Picture1.Visible = False
126             ReversePayment_Click

128         Case 7
130             Picture1.Visible = False
132             Rev_Click

134         Case 8
136             Picture1.Visible = False
138             Co_Click

140         Case 9
142             Picture1.Visible = False
144             viewwedin_Click

146         Case 10
148             Picture1.Visible = False
150             viewlang_Click

152         Case 11
                
                'CloseChildMDI

154             Picture1.Visible = True
                
        End Select

        '<EhFooter>

        TxtLog "Exited Toolbar1_ButtonClick"

        Exit Sub

Toolbar1_ButtonClick_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.Toolbar1_ButtonClick", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       tra_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub tra_Click()

        '<EhHeader>
        On Error GoTo tra_Click_Err

        TxtLog "Entered tra_Click"

        '</EhHeader>

100     Load frm_Trail
102     frm_Trail.Show

        '<EhFooter>

        TxtLog "Exited tra_Click"

        Exit Sub

tra_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.tra_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       users_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub users_Click()

        '<EhHeader>
        On Error GoTo users_Click_Err

        TxtLog "Entered users_Click"

        '</EhHeader>

100     frm_Users.lblUser.Caption = MDIForm1.lblUserName.Caption
102     frm_Users.lblUserlevel.Caption = rsUser!userlevel
104     Load frm_Users

106     frm_Users.Show

        '<EhFooter>

        TxtLog "Exited users_Click"

        Exit Sub

users_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.users_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       viewlang_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub viewlang_Click()

        '<EhHeader>
        On Error GoTo viewlang_Click_Err

        TxtLog "Entered viewlang_Click"

        '</EhHeader>

100     Load frm_CustomerPayment
102     frm_CustomerPayment.Show

        '<EhFooter>

        TxtLog "Exited viewlang_Click"

        Exit Sub

viewlang_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.viewlang_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       viewwedin_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       9/14/2017-8:42:01 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub viewwedin_Click()

        '<EhHeader>
        On Error GoTo viewwedin_Click_Err

        TxtLog "Entered viewwedin_Click"

        '</EhHeader>

100     Load frm_CustomerLoan
102     frm_CustomerLoan.Show

        '<EhFooter>

        TxtLog "Exited viewwedin_Click"

        Exit Sub

viewwedin_Click_Err:
        ErrReport Err.Description, "LendingClient.MDIForm1.viewwedin_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

