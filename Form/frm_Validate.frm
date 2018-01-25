VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Validate 
   Caption         =   "frmValidate"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   6045
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Set Time and Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton btnProceed 
         Caption         =   "&Proceed"
         Height          =   615
         Left            =   3240
         TabIndex        =   5
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton btnSet 
         Caption         =   "&Set"
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5535
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   4560
            Top             =   1080
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   495
            Left            =   1680
            TabIndex        =   3
            Top             =   360
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   873
            _Version        =   393216
            Format          =   265617409
            CurrentDate     =   41869
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Date    :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   4
            Top             =   360
            Width           =   825
         End
      End
   End
   Begin VB.Label lbltime 
      Caption         =   "Label2"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label2"
      Height          =   735
      Left            =   3240
      TabIndex        =   8
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lblusername 
      Caption         =   "Label2"
      Height          =   735
      Left            =   480
      TabIndex        =   7
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label lblUserlevel 
      Caption         =   "Label2"
      Height          =   615
      Left            =   1200
      TabIndex        =   6
      Top             =   3960
      Width           =   2055
   End
End
Attribute VB_Name = "frm_Validate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Validate
'    Project    : Project1
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Private Sub btnProceed_Click()

        '<EhHeader>
        On Error GoTo btnProceed_Click_Err

        TxtLog "Entered btnProceed_Click"

        '</EhHeader>

100     If rsLogtime.State = 1 Then rsLogtime.Close
102     rsLogtime.Open "Select * from tblLogtime"
    
104     If rsLogtime.RecordCount = 0 Then
106         rsLogtime.AddNew
108         rsLogtime!Logdate = DTPicker1.Value
110         rsLogtime.Update
        Else
112         rsLogtime!Logdate = DTPicker1.Value
114         rsLogtime.Update
        End If
        
116     rsTrail.AddNew
118     rsTrail!userlevel = lblUserlevel.Caption
120     rsTrail!UserName = lblUserName.Caption
122     rsTrail!Activity = "Logged In"
124     rsTrail!Time = lblTime.Caption
126     rsTrail!Date = DTPicker1.Value
128     rsTrail!Status = "Log-in"
130     rsTrail.Update
    
132     MDIForm1.lblDate.Caption = rsTrail!Date
134     MDIForm1.lblUserlevel.Caption = rsTrail!userlevel
136     MDIForm1.lblUserName.Caption = rsTrail!UserName

138     If lblUserlevel.Caption = "User" Then
140         MDIForm1.users.Enabled = False
142         Unload Me
144         Load MDIForm1
146         MDIForm1.Show
        Else
148         Unload Me
150         Load MDIForm1
152         MDIForm1.Show
        End If

        '<EhFooter>

        TxtLog "Exited btnProceed_Click"

        Exit Sub

btnProceed_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Validate.btnProceed_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnProceed_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo btnProceed_KeyPress_Err

        TxtLog "Entered btnProceed_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         btnProceed.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited btnProceed_KeyPress"

        Exit Sub

btnProceed_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Validate.btnProceed_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnSet_Click()

        '<EhHeader>
        On Error GoTo btnSet_Click_Err

        TxtLog "Entered btnSet_Click"

        '</EhHeader>

100     If btnSet.Caption = "&Set" Then
102         btnSet.Caption = "%Validate"
104         DTPicker1.Enabled = True
    
        Else
        End If

        '<EhFooter>

        TxtLog "Exited btnSet_Click"

        Exit Sub

btnSet_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Validate.btnSet_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call Logtime
104     Call Trail
        '     Call btnLogin_Click

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Validate.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    DTPicker1.Value = FormatDateTime(Date, vbShortDate)
    lblTime.Caption = Time

End Sub
