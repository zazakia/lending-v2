VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form rep_CollectorList 
   Caption         =   "rep_CollectorList"
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14535
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
   ScaleHeight     =   10350
   ScaleWidth      =   14535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   11895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   20895
      _cx             =   36856
      _cy             =   20981
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
      LocaleID        =   13321
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "rep_CollectorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

        Dim CRApp  As New CRAXDDRT.Application

        Dim Report As New CRAXDDRT.Report

100     Set Report = CRApp.OpenReport(App.Path & "\report\collector list updated.rpt")
102     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
104     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
106     CRViewer.ReportSource = Report
108     CRViewer.viewReport
110     CRViewer.Zoom (100)
112     rep_CollectorList.WindowState = vbMaximized

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.rep_CollectorList.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub
