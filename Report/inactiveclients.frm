VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form rep_inactivecleints_past 
   Caption         =   "rep_inactiveclients"
   ClientHeight    =   10200
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   14145
   LinkTopic       =   "Form2"
   ScaleHeight     =   10200
   ScaleWidth      =   14145
   StartUpPosition =   3  'Windows Default
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   10215
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   14175
      _cx             =   25003
      _cy             =   18018
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
End
Attribute VB_Name = "rep_inactivecleints_past"
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

100     Set Report = CRApp.OpenReport(App.Path & "\report\inactive clients.rpt")
102     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
104     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
        'Report.Database.Tables(1).Location = App.Path & "\JCashdb.mdb"
        'Report.Database.Tables(1).ConnectionProperties("Database Password") = "MyPassword"
106     CRViewer1.ReportSource = Report
        'Report.RecordSelectionFormula = "{tblCustomer.Collector} = 'Mague'"
108     CRViewer1.viewReport
110     CRViewer1.Zoom (1)

112     rep_inactivecleints.WindowState = vbMaximized  'wrong spelling ang clients. hahah.
114     rep_inactivecleints.Refresh

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.rep_inactivecleints_past.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub
