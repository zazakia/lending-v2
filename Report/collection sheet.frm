VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rep_CollectionSheet 
   BackColor       =   &H8000000D&
   Caption         =   "Collection Sheet"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   ScaleHeight     =   10680
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   10695
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   20175
      _cx             =   35586
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
      LocaleID        =   13321
      EnableInteractiveParameterPrompting=   0   'False
   End
   Begin MSComCtl2.DTPicker dtp_cs 
      Height          =   735
      Left            =   5520
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
      Format          =   265617409
      CurrentDate     =   41891
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
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   4020
   End
End
Attribute VB_Name = "rep_CollectionSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : rep_CollectionSheet
'    Project    : Project1
'
'    Description: This form should the costumers of the collectors and the amount the costumers have to pay the collector. This is based by the date.
'
'    Modified   : Kenji, added the date range functionality in the form. By adding to date pickers and editing the selection formula
'                 But it was removed because it was incomplete and errors were found. Completion needed before adition.
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Private Sub RunReport()

        '<EhHeader>
        On Error GoTo RunReport_Err

        TxtLog "Entered RunReport"

        '</EhHeader>

100     Call CollCode
102     Call CollData
        
        Dim CRApp  As New CRAXDDRT.Application

        Dim Report As New CRAXDDRT.Report

104     Set Report = CRApp.OpenReport(App.Path & "\report\collections.rpt")
106     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
108     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
        'Report.Database.Tables(1).Location = App.Path & "\JCashdb.mdb"
        'Report.Database.Tables(1).ConnectionProperties("Database Password") = "MyPassword"
110     CRViewer.ReportSource = Report
        'Report.RecordSelectionFormula = "{tblCustomer.Collector} = 'Mague'"
112     CRViewer.viewReport
114     CRViewer.Zoom (1)
        'Report.RecordSelectionFormula = "{tblLoan.DateRelease}  < #" & dtp_cs & "# and {tblLoan.Status} = 'Good'"
        'Report.RecordSelectionFormula = "{tblLoan.DateRelease}  < #" & dtp_cs & "# and {tblLoan.Status} = 'Good'"
        '{tblLoan.Maturity} >= CurrentDate and
        '{tblCustomer.Balance} <> 0 and
        '{tblLoan.Status} = "Good"

        'i-loop nlang ang mga current collector datas ana na date
        'i declare diri ang mga holders diri. duha man to ka mga holders ang gamit ganina.
        Dim rctr         As Integer

        Dim collectorIDs As String

        Dim switch       As Boolean
        
116     collectorIDs = ""
118     switch = False
        
120     If rsCollCode.State = 1 Then rsCollCode.Close
122     rsCollCode.Open "Select * from tblColl_Code"
        
124     If rsCollCode.RecordCount <> 0 Then
126         rsCollCode.MoveFirst
        
128         For rctr = 0 To (rsCollCode.RecordCount - 1)

130             If rsCollData.State = 1 Then rsCollData.Close
132             rsCollData.Open "Select * from tblColl_Data where Code = " & _
                        rsCollCode!code & " and DateEmployed <= #" & dtp_cs & _
                        "# order by DateEmployed DESC"
                
134             If rsCollData.RecordCount <> 0 Then
136                 switch = True
                    
138                 If collectorIDs = "" Then
140                     collectorIDs = " and  {tblColl_Data.ID} in [" & rsCollData!ID
                    Else
142                     collectorIDs = collectorIDs & ", " & rsCollData!ID
                    End If
                End If
                
144             rsCollCode.MoveNext
146         Next rctr
                       
            'ngari na after sa loop i check kung ni sud ba siya kausa? nya butang dayon sa closing bracket
148         If switch Then collectorIDs = collectorIDs & "]"
            
        End If
                
        'MsgBox (collectorIDs)
        'Edited by Brayan 1/5/18 Friday 12pm. To combine all clients active and inactive
150     Report.RecordSelectionFormula = _
                "{tblLoan.Status} = 'Good' and {tblCustomer.Balance} <> 0 "
        '152     Report.RecordSelectionFormula = _
        '                "{tblLoan.Status} = 'Good' and {tblCustomer.Balance} <> 0 and {tblLoan.Maturity} >= #" _
        '                & dtp_cs & "#" & collectorIDs
152     rep_CollectionSheet.Refresh

        '<EhFooter>

        TxtLog "Exited RunReport"

        Exit Sub

RunReport_Err:
        ErrReport Err.Description, "LendingClient.rep_CollectionSheet.RunReport", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub dtp_cs_Change()

        '<EhHeader>
        On Error GoTo dtp_cs_Change_Err

        TxtLog "Entered dtp_cs_Change"

        '</EhHeader>

100     RunReport

        '<EhFooter>

        TxtLog "Exited dtp_cs_Change"

        Exit Sub

dtp_cs_Change_Err:
        ErrReport Err.Description, "LendingClient.rep_CollectionSheet.dtp_cs_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     dtp_cs = DateValue(Now) + 1
102     RunReport

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.rep_CollectionSheet.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

