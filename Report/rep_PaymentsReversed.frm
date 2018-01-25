VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rep_PaymentsReversed 
   Caption         =   "Payments Reversed"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15210
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
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtp_From 
      Height          =   735
      Left            =   1680
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
      Height          =   10200
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   15255
      _cx             =   26908
      _cy             =   17992
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
   Begin MSComCtl2.DTPicker dtp_To 
      Height          =   735
      Left            =   6360
      TabIndex        =   4
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
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   795
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "rep_PaymentsReversed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub viewReport()

        '<EhHeader>
        On Error GoTo viewReport_Err

        TxtLog "Entered viewReport"

        '</EhHeader>

        Dim CRApp  As New CRAXDRT.Application

        Dim Report As New CRAXDRT.Report

100     Set Report = CRApp.OpenReport(App.Path & "\report\pdr.rpt")
102     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
104     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
106     CRViewer1.ReportSource = Report
108     CRViewer1.viewReport
110     CRViewer1.Zoom (1)

        Dim rctr         As Integer

        Dim collectorIDs As String

        Dim switch       As Boolean
        
112     collectorIDs = ""
114     switch = False
        
116     If rsCollCode.State = 1 Then rsCollCode.Close
118     rsCollCode.Open "Select * from tblColl_Code"
        
120     If rsCollCode.RecordCount <> 0 Then
122         rsCollCode.MoveFirst
        
124         For rctr = 0 To (rsCollCode.RecordCount - 1)

126             If rsCollData.State = 1 Then rsCollData.Close
128             rsCollData.Open "Select * from tblColl_Data where Code = '" & _
                        rsCollCode!code & "' and DateEmployed <= #" & dtp_From & _
                        "# order by DateEmployed DESC"
                
130             If rsCollData.RecordCount <> 0 Then
132                 switch = True
                    
134                 If collectorIDs = "" Then
136                     collectorIDs = " and  {tblColl_Data.ID} in [" & rsCollData!ID
                    Else
138                     collectorIDs = collectorIDs & ", " & rsCollData!ID
                    End If
                End If
                
140             rsCollCode.MoveNext
142         Next rctr
                       
            'ngari na after sa loop i check kung ni sud ba siya kausa? nya butang dayon sa closing bracket
144         If switch Then collectorIDs = collectorIDs & "]"
            
        End If

146     Report.RecordSelectionFormula = "({tblPayment.Date} in Date(#" & dtp_From & _
                "#) to Date(#" & dtp_To & "#)) and ({tblPayment.Status} = 'Reversing')" _
                & collectorIDs
148     rep_PaymentsReversed.WindowState = vbNormal
150     rep_PaymentsReversed.Refresh

        '<EhFooter>

        TxtLog "Exited viewReport"

        Exit Sub

viewReport_Err:
        ErrReport Err.Description, "LendingClient.rep_PaymentsReversed.viewReport", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub dtp_From_change()

        '<EhHeader>
        On Error GoTo dtp_From_change_Err

        TxtLog "Entered dtp_From_change"

        '</EhHeader>

100     dtp_To.Value = dtp_From.Value
102     Call viewReport

        '<EhFooter>

        TxtLog "Exited dtp_From_change"

        Exit Sub

dtp_From_change_Err:
        ErrReport Err.Description, _
                "LendingClient.rep_PaymentsReversed.dtp_From_change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub dtp_To_Change()

        '<EhHeader>
        On Error GoTo dtp_To_Change_Err

        TxtLog "Entered dtp_To_Change"

        '</EhHeader>

100     Call viewReport

        '<EhFooter>

        TxtLog "Exited dtp_To_Change"

        Exit Sub

dtp_To_Change_Err:
        ErrReport Err.Description, "LendingClient.rep_PaymentsReversed.dtp_To_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call CollCode
102     Call CollData
    
104     dtp_From.Value = DateTime.Now
106     dtp_To.Value = DateTime.Now
108     Call viewReport

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.rep_PaymentsReversed.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

