VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rep_month_released 
   Caption         =   "Form2"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4950
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   260177921
      CurrentDate     =   41913
   End
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer crystal_monthly_released 
      Height          =   12615
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   14415
      _cx             =   25426
      _cy             =   22251
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   260177921
      CurrentDate     =   41913
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   690
   End
End
Attribute VB_Name = "rep_month_released"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub dtpTo_Change()
        '<EhHeader>
        On Error GoTo dtpTo_Change_Err
        TxtLog "Entered dtpTo_Change"
        '</EhHeader>


        Dim CRApp  As New CRAXDDRT.Application

        Dim Report As New CRAXDDRT.Report
   
100     Set Report = CRApp.OpenReport(App.Path & "\report\monthly_released.rpt ")
102     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
104     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
106     crystal_monthly_released.ReportSource = Report
108     crystal_monthly_released.viewReport
110     crystal_monthly_released.Zoom (1)
112     Report.RecordSelectionFormula = "({tblLoan.DateRelease} in Date(#" & dtpFrom & _
                "#) to Date(#" & dtpTo & _
                "#)) and ({tblLoan.Status} = 'Good' or {tblLoan.Status} = 'Full Paid')"
114     rep_month_released.WindowState = vbMaximized
116     rep_month_released.Refresh


        '<EhFooter>
        TxtLog "Exited dtpTo_Change"
        Exit Sub

dtpTo_Change_Err:
        ErrReport Err.Description, "LendingClient.rep_month_released.dtpTo_Change", Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dtpFrom_Change()
        '<EhHeader>
        On Error GoTo dtpFrom_Change_Err
        TxtLog "Entered dtpFrom_Change"
        '</EhHeader>

        Dim CRApp  As New CRAXDDRT.Application

        Dim Report As New CRAXDDRT.Report

100     Set Report = CRApp.OpenReport(App.Path & "\report\monthly_released.rpt ")
102     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
104     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
106     crystal_monthly_released.ReportSource = Report
108     crystal_monthly_released.viewReport
110     crystal_monthly_released.Zoom (1)
112     Report.RecordSelectionFormula = "({tblLoan.DateRelease} in Date(#" & dtpFrom & _
                "#) to Date(#" & dtpTo & _
                "#)) and ({tblLoan.Status} = 'Good' or {tblLoan.Status} = 'Full Paid')"
114     rep_month_released.WindowState = vbMaximized
116     rep_month_released.Refresh


        '<EhFooter>
        TxtLog "Exited dtpFrom_Change"
        Exit Sub

dtpFrom_Change_Err:
        ErrReport Err.Description, "LendingClient.rep_month_released.dtpFrom_Change", Erl
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

100     dtpFrom.Value = DateTime.Now
102     dtpTo.Value = DateTime.Now
104     Set Report = CRApp.OpenReport(App.Path & "\report\monthly_released.rpt ")
106     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
108     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
110     crystal_monthly_released.ReportSource = Report
112     crystal_monthly_released.viewReport
114     crystal_monthly_released.Zoom (1)
116     Report.RecordSelectionFormula = "({tblLoan.DateRelease} in Date(#" & dtpFrom & _
                "#) to Date(#" & dtpTo & _
                "#)) and ({tblLoan.Status} = 'Good' or {tblLoan.Status} = 'Full Paid')"
118     rep_month_released.WindowState = vbMaximized
120     rep_month_released.Refresh


        '<EhFooter>
        TxtLog "Exited Form_Load"
        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.rep_month_released.Form_Load", Erl
        Resume Next
        '</EhFooter>
End Sub
