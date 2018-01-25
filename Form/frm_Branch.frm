VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Branch 
   AutoRedraw      =   -1  'True
   Caption         =   "frmBranch"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   12840
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5175
      Left            =   360
      TabIndex        =   1
      Top             =   4800
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9128
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "frmBranch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11655
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      Begin VB.TextBox txtAprvPos 
         Height          =   375
         Left            =   8280
         TabIndex        =   23
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtApprovedBy 
         Height          =   375
         Left            =   8280
         TabIndex        =   22
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtChkPos 
         Height          =   405
         Left            =   8280
         TabIndex        =   21
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtCheckBy 
         Height          =   405
         Left            =   2520
         TabIndex        =   20
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox txtRcvPos 
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox txtReceivedby 
         Height          =   375
         Left            =   2520
         TabIndex        =   18
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtTodayDate 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   7200
         TabIndex        =   14
         Top             =   240
         Width           =   2775
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   4560
         Top             =   480
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   5280
         TabIndex        =   10
         Top             =   9960
         Width           =   1695
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2640
         TabIndex        =   9
         Top             =   9960
         Width           =   1695
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   9960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtLocation 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtBranchName 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   4200
         Width           =   3255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Approved Position  :"
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
         Left            =   6000
         TabIndex        =   29
         Top             =   2160
         Width           =   2100
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Approved By            :"
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
         Left            =   6000
         TabIndex        =   28
         Top             =   1680
         Width           =   2130
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Check Position        :"
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
         Left            =   6000
         TabIndex        =   27
         Top             =   1200
         Width           =   2115
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Check By                   :"
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
         Left            =   120
         TabIndex        =   26
         Top             =   3120
         Width           =   2205
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Recieved  Position   :"
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
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   2190
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Recieved By             :"
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
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   2160
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Date: "
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
         Left            =   6480
         TabIndex        =   15
         Top             =   240
         Width           =   645
      End
      Begin VB.Line Line1 
         X1              =   -120
         X2              =   12000
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Branch Location      :"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   2160
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Branch Name           :"
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
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search here:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   4200
         Width           =   1365
      End
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   8760
      TabIndex        =   17
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label lblBranchId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   8760
      TabIndex        =   16
      Top             =   2760
      Width           =   45
   End
   Begin VB.Label lblTime 
      Caption         =   "Label5"
      Height          =   255
      Left            =   12480
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUserlevel 
      Caption         =   "Label5"
      Height          =   255
      Left            =   12480
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8520
      TabIndex        =   11
      Top             =   2040
      Width           =   45
   End
End
Attribute VB_Name = "frm_Branch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Branch
'    Project    : Project1
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Private Sub btnAdd_Click()

        '<EhHeader>
        On Error GoTo btnAdd_Click_Err

        TxtLog "Entered btnAdd_Click"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then
102         btnAdd.Caption = "&Save"
104         txtBranchName.Enabled = True
106         txtLocation.Enabled = True
108         txtBranchName.SetFocus
        Else

110         If txtBranchName.Text = "" Or txtLocation.Text = "" Then
112             MsgBox "All fields are required.", vbInformation, _
                        "Webplus Lending Corporation"
114             txtBranchName.SetFocus
            Else

116             If MsgBox("Are you sure you want  to create New Branch Record?", _
                        vbQuestion + vbYesNo) = vbYes Then
    
118                 If rsBranch.State = 1 Then rsBranch.Close
120                 rsBranch.Open "Select * from tblBranch where BranchName = '" & _
                            txtBranchName.Text & "'"

122                 If rsBranch.RecordCount <> 0 Then
124                     MsgBox "Record Already exist.", vbInformation, _
                                "Webplus Lending Corporation"
126                     Call Branch
                    Else

128                     With rsBranch
130                         .AddNew
132                         !BranchName = txtBranchName.Text
134                         !BranchLocation = txtLocation.Text
136                         !RecievedBy = txtReceivedby.Text
138                         !RcvPos = txtRcvPos.Text
140                         !CheckBy = txtCheckBy.Text
142                         !ChkPos = txtChkPos.Text
144                         !ApprovedBy = txtApprovedBy.Text
146                         !AprvPos = txtAprvPos.Text
148                         .Update
                
                        End With

150                     If rsTrail.State = 1 Then rsTrail.Close
152                     rsTrail.Open "Select * from tblTrail where Username = '" & _
                                lblUser.Caption & "'"

154                     With rsTrail
156                         .AddNew
158                         !UserName = lblUser.Caption
160                         !userlevel = lblUserlevel.Caption
162                         !Activity = "Add New Income Record"
164                         !Time = lblTime.Caption
166                         !Date = txtTodayDate.Text
168                         .Update
                        End With
            
170                     MsgBox "Record successfully Added", vbInformation, _
                                "Webplus Lending Corporation"
172                     Unload Me
174                     Me.Show
        
                    End If
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnAdd_Click"

        Exit Sub

btnAdd_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Branch.btnAdd_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnClose_Click()

        '<EhHeader>
        On Error GoTo btnClose_Click_Err

        TxtLog "Entered btnClose_Click"

        '</EhHeader>

100     If btnClose.Caption = "&Close" Then
102         Unload Me
        Else
104         btnClose.Caption = "&Close"
106         btnEdit.Enabled = False
108         btnAdd.Enabled = True
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Branch.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEdit_Click()

        '<EhHeader>
        On Error GoTo btnEdit_Click_Err

        TxtLog "Entered btnEdit_Click"

        '</EhHeader>

100     If btnEdit.Caption = "&Edit" Then
102         btnEdit.Caption = "&Update"
104         txtBranchName.Enabled = True
106         txtLocation.Enabled = True
108         txtBranchName.SetFocus
        Else

110         If rsBranch.State = 1 Then rsBranch.Close
112         rsBranch.Open "Select * from tblBranch where BranchName = '" & _
                    lblName.Caption & "'"

114         If rsBranch.RecordCount = 0 Then
116             MsgBox "No Record Found", vbInformation
            Else

118             With rsBranch
120                 !BranchName = txtBranchName.Text
122                 !BranchLocation = txtLocation.Text
124                 !RecievedBy = txtReceivedby.Text
126                 !RcvPos = txtRcvPos.Text
128                 !CheckBy = txtCheckBy.Text
130                 !ChkPos = txtChkPos.Text
132                 !ApprovedBy = txtApprovedBy.Text
134                 !AprvPos = txtAprvPos.Text
136                 .Update
                    
                End With
                
138             MsgBox "Branch record successfully updated", vbInformation
140             Unload Me
142             Me.Show
            End If
    
        End If

        '<EhFooter>

        TxtLog "Exited btnEdit_Click"

        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Branch.btnEdit_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        TxtLog "Entered DataGrid1_Click"

        '</EhHeader>

100     If rsBranch.RecordCount = 0 Then
        Else
102         txtBranchName.Text = rsBranch!BranchName
104         lblName.Caption = rsBranch!BranchName
106         txtLocation.Text = rsBranch!BranchLocation
108         lblBranchId.Caption = rsBranch!BranchID
110         txtReceivedby.Text = rsBranch!RecievedBy
112         txtRcvPos.Text = rsBranch!RcvPos
114         txtCheckBy.Text = rsBranch!CheckBy
116         txtChkPos.Text = rsBranch!ChkPos
118         txtApprovedBy.Text = rsBranch!ApprovedBy
120         txtAprvPos.Text = rsBranch!AprvPos

122         btnAdd.Enabled = False
124         btnClose.Caption = "&Cancel"
126         btnEdit.Enabled = True
        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_Click"

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Branch.DataGrid1_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DataGrid1_KeyPress_Err

        TxtLog "Entered DataGrid1_KeyPress"

        '</EhHeader>

100     KeyAscii = 0

        '<EhFooter>

        TxtLog "Exited DataGrid1_KeyPress"

        Exit Sub

DataGrid1_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Branch.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call Branch
104     Call Trail

106     Set DataGrid1.DataSource = rsBranch

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Branch.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    lblTime.Caption = Time
    txtTodayDate.Text = Date
End Sub

Private Sub txtBranchName_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtBranchName_KeyPress_Err

        TxtLog "Entered txtBranchName_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         txtLocation.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtBranchName_KeyPress"

        Exit Sub

txtBranchName_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Branch.txtBranchName_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsBranch.State = 1 Then rsBranch.Close
102     rsBranch.Open "Select * from tblBranch where BranchName like '" & _
                txtSearch.Text & "%'"

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Branch.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

