VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rep_cpr 
   Caption         =   "rep_cpr"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13140
   ScaleHeight     =   10935
   ScaleWidth      =   13140
   StartUpPosition =   3  'Windows Default
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   9615
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   12495
      _cx             =   22040
      _cy             =   16960
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
   Begin MSComCtl2.DTPicker dtp_cpr 
      Height          =   735
      Left            =   4440
      TabIndex        =   1
      Top             =   120
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
      Format          =   96600065
      CurrentDate     =   41891
   End
   Begin VB.Label testlabel 
      Caption         =   "Label1"
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblSelectDate 
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4020
   End
End
Attribute VB_Name = "rep_cpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dtp_cpr_Change()

        '<EhHeader>
        On Error GoTo dtp_cpr_Change_Err

        TxtLog "Entered dtp_cpr_Change"

        '</EhHeader>

100     Form_Load

        '<EhFooter>

        TxtLog "Exited dtp_cpr_Change"

        Exit Sub

dtp_cpr_Change_Err:
        ErrReport Err.Description, "LendingClient.rep_cpr.dtp_cpr_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     testlabel.Caption = rep_cprForm.kita.Caption

        Dim CRApp  As New CRAXDDRT.Application

        Dim Report As New CRAXDDRT.Report

102     Set Report = CRApp.OpenReport(App.Path & "\report\CPR.rpt ")
        
104     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
106     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
108     CRViewer.ReportSource = Report
110     CRViewer.viewReport
112     CRViewer.Zoom (1)
        
114     CRViewer.DisplayTabs = False
116     CRViewer.EnableGroupTree = False
118     CRViewer.EnableCloseButton = False
120     CRViewer.EnableStopButton = False
122     CRViewer.EnableExportButton = False
124     Report.RecordSelectionFormula = "{tblCustomer.Code} = " & testlabel & ""
126     rep_cpr.WindowState = vbMaximized
128     rep_cpr.Refresh

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.rep_cpr.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub
