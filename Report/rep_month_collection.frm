VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rep_month_collection 
   Caption         =   "Monthly Collection"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   14400
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   14400
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtp_dp 
      Height          =   735
      Left            =   8040
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
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
      Format          =   266010625
      CurrentDate     =   41891
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   735
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
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
      Format          =   266010625
      CurrentDate     =   41891
   End
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   9975
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   13830
      _cx             =   24395
      _cy             =   17595
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   13321
      EnableInteractiveParameterPrompting=   0   'False
   End
   Begin VB.Label L 
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label L 
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   6960
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "rep_month_collection"
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
        ErrReport Err.Description, "LendingClient.rep_month_collection.dtp_dp_change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DTPicker1_Change()

        '<EhHeader>
        On Error GoTo DTPicker1_Change_Err

        TxtLog "Entered DTPicker1_Change"

        '</EhHeader>

100     dtp_dp.Value = DTPicker1.Value

102     Form_Load

        '<EhFooter>

        TxtLog "Exited DTPicker1_Change"

        Exit Sub

DTPicker1_Change_Err:
        ErrReport Err.Description, _
                "LendingClient.rep_month_collection.DTPicker1_Change", Erl

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

100     Set Report = CRApp.OpenReport(App.Path & "\report\monthly_collection.rpt ")
102     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
104     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
106     CRViewer1.ReportSource = Report
        
108     CRViewer1.viewReport
110     CRViewer1.Zoom (1)
112     Report.RecordSelectionFormula = "({tblPayment.Date} in Date(#" & DTPicker1 & _
                "#) to Date(#" & dtp_dp & _
                "#)) and ({tblPayment.Status} = 'Good' or {tblPayment.Status} = 'Full Paid')"
114     rep_month_collection.WindowState = vbMaximized
116     rep_month_collection.Refresh

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.rep_month_collection.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

