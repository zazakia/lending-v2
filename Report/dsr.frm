VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rep_DailySalesReport 
   Caption         =   "rep_DailySalesReport"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14010
   LinkTopic       =   "Form2"
   ScaleHeight     =   10395
   ScaleWidth      =   14010
   StartUpPosition =   2  'CenterScreen
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   9705
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   13935
      _cx             =   24580
      _cy             =   17119
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
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   735
      Left            =   6240
      TabIndex        =   3
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
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
      Format          =   265551873
      CurrentDate     =   41891
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   735
      Left            =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
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
      Format          =   265551873
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
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   705
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
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "rep_DailySalesReport"
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
        
        Dim CRApp  As New CRAXDDRT.Application

        Dim Report As New CRAXDDRT.Report

100     Set Report = CRApp.OpenReport(App.Path & "\report\DSR LOANS AS MAIN.rpt ")
102     Report.Database.Tables(1).Location = App.Path & "\db\JCashdb.mdb"
104     Report.Database.Tables(1).SetSessionInfo "", Chr$(10) & "kim123"
106     CRViewer.ReportSource = Report
108     CRViewer.viewReport
110     CRViewer.Zoom (1)
    
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
128             rsCollData.Open "Select * from tblColl_Data where Code = " & _
                        rsCollCode!code & " and DateEmployed <= #" & dtpFrom & _
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
    
146     Report.RecordSelectionFormula = "({tblLoan.DateRelease} in Date(#" & dtpFrom & _
                "#) to Date(#" & dtpTo & _
                "#)) and ({tblLoan.Status} = 'Good' or {tblLoan.Status} = 'Full Paid')"
        '& collectorIDs
        '"({tblPayment.Date} in Date(#" & dtpFrom & "#) to Date(#" & dtpTo & "#)) and ({tblPayment.Status} = 'Good' or {tblPayment.Status} = 'Full Paid')"
148     rep_DailySalesReport.WindowState = vbNormal
150     rep_DailySalesReport.Refresh

        '<EhFooter>

        TxtLog "Exited viewReport"

        Exit Sub

viewReport_Err:
        ErrReport Err.Description, "LendingClient.rep_DailySalesReport.viewReport", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub dtpFrom_Change()

        '<EhHeader>
        On Error GoTo dtpFrom_Change_Err

        TxtLog "Entered dtpFrom_Change"

        '</EhHeader>

100     dtpTo.Value = dtpFrom.Value
        'benjamin sumilhig
102     Call enter_date
104     Call viewReport

        '<EhFooter>

        TxtLog "Exited dtpFrom_Change"

        Exit Sub

dtpFrom_Change_Err:
        ErrReport Err.Description, "LendingClient.rep_DailySalesReport.dtpFrom_Change", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub dtpTo_Change()

        '<EhHeader>
        On Error GoTo dtpTo_Change_Err

        TxtLog "Entered dtpTo_Change"

        '</EhHeader>

100     Call viewReport

        '<EhFooter>

        TxtLog "Exited dtpTo_Change"

        Exit Sub

dtpTo_Change_Err:
        ErrReport Err.Description, "LendingClient.rep_DailySalesReport.dtpTo_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

'Benjamin Sumilhig
'mo ni ang mo enter sa date sa tbl loan if walay me loan anang adlawa
'para nay report sa collection sa DCR report
Private Sub enter_date()

        '<EhHeader>
        On Error GoTo enter_date_Err

        TxtLog "Entered enter_date"

        '</EhHeader>

100     If rsPayment.State = 1 Then rsPayment.Close
102     rsPayment.Open "SELECT * FROM tblPayment WHERE Date = #" & Format$( _
                dtpFrom.Value, "mm/dd/yy") & "#"

104     If rsPayment.RecordCount > 0 Then
106         If rsLoan.State = 1 Then rsLoan.Close
108         rsLoan.Open "SELECT * FROM tblLoan WHERE DateRelease = #" & Format$( _
                    dtpFrom.Value, "mm/dd/yy") & "# And Status = 'Good'"

110         If rsLoan.RecordCount = 0 Then

                'MsgBox "=0"
112             With rsLoan
114                 .AddNew
116                 !Collector = " "
118                 !code = " "
120                 !Customer = " "
122                 !firstname = " "
124                 !principal = 0
126                 !total = 0
128                 !LoanDate = Format$(dtpFrom.Value, "m/d/yy")
130                 !DateRelease = Format$(dtpFrom.Value, "m/d/yy")
132                 !Maturity = Format$(dtpFrom.Value, "m/d/yy")
134                 !Status = "Good"
136                 !FireInsurance = 0
138                 !CollectorCharge = 0
140                 !delivery = 0
142                 !collection = 0
144                 !Servicefee = 0
146                 !Balance = 0
148                 !Penalty = 0
150                 !Passbook = 0
152                 !TotalAmortization = 0
154                 !TotalCharges = 0
156                 !LoanTotal = 0
158                 !CollectorCode = " "
160                 !CollectorFname = " "
162                 !User = " "
164                 !LoanStatus = "Good"
166                 !TotalPayment = 0
168                 !NotPosted = 0
170                 !OverToCus = 0
172                 .Update
                End With

            Else
                'Blanko lang ning else
                'MsgBox ">0"
            End If
        End If

        '<EhFooter>

        TxtLog "Exited enter_date"

        Exit Sub

enter_date_Err:
        ErrReport Err.Description, "LendingClient.rep_DailySalesReport.enter_date", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call Loan
102     Call payment
104     Call CollCode
106     Call CollData
108     dtpFrom.Value = DateTime.Now
110     dtpTo.Value = DateTime.Now
        'benjamin sumilhig
112     Call enter_date
114     Call viewReport

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.rep_DailySalesReport.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

