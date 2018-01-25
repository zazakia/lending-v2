VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rep_dailyPayment 
   Caption         =   "Daily Payment Report"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CrystalActiveXReportViewer5 
      Height          =   10695
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   15135
      _cx             =   26696
      _cy             =   18865
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
      EnableInteractiveParameterPrompting=   0   'False
   End
   Begin MSComCtl2.DTPicker dtp_dp 
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   264765441
      CurrentDate     =   41891
   End
   Begin VB.Label lblSelectDate2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   4020
   End
End
Attribute VB_Name = "rep_dailyPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dtp_dp_change()

        '<EhHeader>
        On Error GoTo dtp_dp_change_Err

        TxtLog "Entered dtp_dp_change"

        '</EhHeader>

100     Form_Load

        '<EhFooter>

        TxtLog "Exited dtp_dp_change"

        Exit Sub

dtp_dp_change_Err:
        ErrReport Err.Description, "LendingClient.rep_dailyPayment.dtp_dp_change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

        Dim CRApp  As New CRAXDDRT.Application

        Dim Report As New CRAXDDRT.Report

100     Set Report = CRApp.OpenReport(App.Path & "\report\dailypayments.rpt ")
102     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
104     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
106     CrystalActiveXReportViewer5.ReportSource = Report
        
108     CrystalActiveXReportViewer5.viewReport
110     CrystalActiveXReportViewer5.Zoom (1)
112     Report.RecordSelectionFormula = "{tblPayment.DateRelease}  = #" & dtp_dp & _
                "# and {tblPayment.Status} = 'Good'"
        
        'isnull({tblLoan.LoanStatus}) = true"
114     rep_DailySalesReport.WindowState = vbMaximized
116     rep_DailySalesReport.Refresh

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.rep_dailyPayment.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

