VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Deposit 
   Caption         =   "frmDeposit"
   ClientHeight    =   10035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Other Income and Deposits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   20055
      Begin VB.TextBox txtUser 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   31
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   600
         Left            =   5160
         TabIndex        =   30
         Top             =   9240
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtFilter 
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   530710529
         CurrentDate     =   41897
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   600
         Left            =   2640
         TabIndex        =   24
         Top             =   9240
         Width           =   1815
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   4200
         Width           =   4815
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   7560
         TabIndex        =   11
         Top             =   9240
         Width           =   1815
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   9240
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4455
         Left            =   120
         TabIndex        =   2
         Top             =   4680
         Width           =   19095
         _ExtentX        =   33681
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   -1  'True
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
      Begin VB.Frame Frame2 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   12255
         Begin VB.TextBox txtAmount 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   36
            Text            =   "0"
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox txtDateEncoded 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   405
            Left            =   6720
            TabIndex        =   26
            Top             =   240
            Width           =   2415
         End
         Begin VB.TextBox txtParticular 
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   22
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox txtID 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   1800
            TabIndex        =   16
            Top             =   240
            Width           =   2775
         End
         Begin VB.ComboBox cbAccount 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            TabIndex        =   15
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   6720
            TabIndex        =   4
            Text            =   "0"
            Top             =   2280
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   4800
            Top             =   240
         End
         Begin MSComCtl2.DTPicker dtDate 
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Top             =   840
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   530776065
            CurrentDate     =   41873
         End
         Begin VB.Label lblAuDep 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   11280
            TabIndex        =   33
            Top             =   240
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date :"
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
            TabIndex        =   25
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "ID                  :"
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
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Particular     :"
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
            Left            =   240
            TabIndex        =   9
            Top             =   2280
            Width           =   1350
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Account       :"
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
            Left            =   240
            TabIndex        =   8
            Top             =   1320
            Width           =   1365
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Amount        :"
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
            Left            =   240
            TabIndex        =   7
            Top             =   1800
            Width           =   1380
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Date             :"
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
            Left            =   240
            TabIndex        =   6
            Top             =   840
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total       :"
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
            Left            =   5640
            TabIndex        =   5
            Top             =   2280
            Visible         =   0   'False
            Width           =   1005
         End
      End
      Begin VB.Label TREWER 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   2160
         TabIndex        =   35
         Top             =   3720
         Width           =   90
      End
      Begin VB.Label lblAWEQWE 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH USING ACCOUNT NAME , PARTICULARS , OR DATE"
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
         Left            =   2280
         TabIndex        =   34
         Top             =   3720
         Width           =   6945
      End
      Begin VB.Label lblUseroooo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
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
         Left            =   6120
         TabIndex        =   32
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblFilterDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Date:"
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
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Search Here    :"
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
         Left            =   240
         TabIndex        =   13
         Top             =   4200
         Width           =   1650
      End
   End
   Begin VB.Label lblFillDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   9840
      TabIndex        =   29
      Top             =   2640
      Width           =   45
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ppppp"
      Height          =   195
      Left            =   20280
      TabIndex        =   23
      Top             =   1560
      Width           =   450
   End
   Begin VB.Label lblCash 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   10080
      TabIndex        =   21
      Top             =   5400
      Width           =   45
   End
   Begin VB.Label lblUserlevel 
      Caption         =   "Label9"
      Height          =   135
      Left            =   20280
      TabIndex        =   20
      Top             =   3480
      Width           =   135
   End
   Begin VB.Label lblUser 
      Caption         =   "Label9"
      Height          =   135
      Left            =   20280
      TabIndex        =   19
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label lblTime 
      Caption         =   "Label9"
      Height          =   255
      Left            =   20280
      TabIndex        =   18
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "lblType"
      Height          =   375
      Left            =   20280
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
End
Attribute VB_Name = "frm_Deposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Deposit
'    Project    : Project1
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Sub addItems()

        '<EhHeader>
        On Error GoTo addItems_Err

        TxtLog "Entered addItems"

        '</EhHeader>

100     If rsChart.State = 1 Then rsChart.Close
102     rsChart.Open _
                "Select distinct AccountName from tblChartOfAccounts  where Type = '" & _
                "Income" & "'"
104     cbAccount.Clear

106     If rsChart.RecordCount <> 0 Then

108         Do While Not rsChart.EOF
110             cbAccount.AddItem rsChart!AccountName
112             rsChart.MoveNext
            Loop

        End If

        '<EhFooter>

        TxtLog "Exited addItems"

        Exit Sub

addItems_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.addItems", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub autoNumber()

        '<EhHeader>
        On Error GoTo autoNumber_Err

        TxtLog "Entered autoNumber"

        '</EhHeader>

100     If rsDeposit.RecordCount = 0 Then
102         txtID.Text = "D1"
            '        ElseIf rsDeposit!DepositID >= "D999" Then
            '           ' rsDeposit.MoveLast
            '            txtID.Text = "D" & Format(Right(rsDeposit!DepositID, 5) + 1, "00000")
            '        ElseIf Len(rsDeposit!DepositID) = 4 Then
            '            rsDeposit.MoveLast
            '            txtID.Text = "D" & Format(Right(rsDeposit!DepositID, 5) + 1, "00000")
        Else
            '106         rsDeposit.MoveLast
            '108         txtID.Text = "D" & Format(Right(rsDeposit!DepositID, 3) + 1, "000")
104         txtID.Text = "D" & rsDeposit!ID + 1
106         lblAuDep.Caption = txtID.Text
        End If

        '<EhFooter>

        TxtLog "Exited autoNumber"

        Exit Sub

autoNumber_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.autoNumber", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnAdd_Click()

        '<EhHeader>
        On Error GoTo btnAdd_Click_Err

        TxtLog "Entered btnAdd_Click"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then
102         dtDate.Enabled = True
104         btnAdd.Caption = "&Save"
106         btnClose.Caption = "&Cancel"
108         DataGrid1.Enabled = False
110         cbAccount.Enabled = True
112         txtAmount.Enabled = True
114         txtParticular.Enabled = True
116         cbAccount.SetFocus
118     ElseIf btnAdd.Caption = "&Save" Then

120         If cbAccount.Text = "" Then
122             MsgBox "Account name is required.Please don't leave it blank", _
                        vbInformation
124             cbAccount.SetFocus
126         ElseIf txtAmount.Text = "0" Then
128             MsgBox "Amount should not be zero.", vbInformation
130         ElseIf txtParticular.Text = "" Then
132             MsgBox "Particulars should not  be blank.", vbInformation, _
                        "Webplus Lending Corporation"
134             txtParticular.SetFocus
            Else

136             If MsgBox("Are you sure you Add this Income?", vbQuestion + vbYesNo) = _
                        vbYes Then
138                 txtTotal.Text = txtAmount.Text

                    'Benjamin Sumilhig
140                 If rsLoan.State = 1 Then rsLoan.Close
142                 rsLoan.Open "SELECT * FROM tblLoan WHERE DateRelease = #" & Format$( _
                            dtDate.Value, "mm/dd/yy") & "# And Status = 'Good'"

144                 If rsLoan.RecordCount = 0 Then

                        'MsgBox "=0"
146                     With rsLoan
148                         .AddNew
150                         !Collector = " "
152                         !code = " "
154                         !Customer = " "
156                         !firstname = " "
158                         !principal = 0
160                         !total = 0
162                         !LoanDate = Format$(dtDate.Value, "m/d/yy")
164                         !DateRelease = Format$(dtDate.Value, "m/d/yy")
166                         !Maturity = Format$(dtDate.Value, "m/d/yy")
168                         !Status = "Good"
170                         !FireInsurance = 0
172                         !CollectorCharge = 0
174                         !delivery = 0
176                         !collection = 0
178                         !Servicefee = 0
180                         !Balance = 0
182                         !Penalty = 0
184                         !Passbook = 0
186                         !TotalAmortization = 0
188                         !TotalCharges = 0
190                         !LoanTotal = 0
192                         !CollectorCode = " "
194                         !CollectorFname = " "
196                         !User = " "
198                         !LoanStatus = "Good"
200                         !TotalPayment = 0
202                         !NotPosted = 0
204                         !OverToCus = 0
206                         .Update
                        End With

                    Else
                        'Blanko lang ning else
                        'MsgBox ">0"
                    End If

                    'End of Benjamin Sumilhig
208                 If rsDeposit.State = 1 Then rsDeposit.Close
210                 rsDeposit.Open "Select * from tblDeposit where DepositID = '" & _
                            txtID.Text & "'"

                    'Print rsDeposit.RecordCount
                    'Print rsDeposit!DepositID

212                 If rsDeposit.RecordCount <> 0 Then
                    Else

214                     With rsDeposit
216                         .AddNew
218                         Print !ID
220                         !DepositID = txtID.Text
222                         !Date = dtDate.Value
224                         !Amount = Format$(Val(txtAmount.Text), "###,###,##0.00")
226                         !AccountName = Trim$(cbAccount.Text)
228                         !Particular = Trim$(txtParticular.Text)
230                         !DateEncoded = txtDateEncoded.Text
232                         !total = Format$(Val(txtTotal.Text), "###,###,##0.00")
234                         !User = txtUser.Text
236                         .Update
                        End With
                                
238                     If rsTrail.State = 1 Then rsTrail.Close
240                     rsTrail.Open "Select * from tblTrail where Username = '" & _
                                txtUser.Text & "'"

242                     With rsTrail
244                         .AddNew
246                         !UserName = txtUser.Text
248                         !userlevel = lblUserlevel.Caption
250                         !Activity = "Add New Income Record"
252                         !Time = lblTime.Caption
254                         !Date = dtDate.Value
256                         .Update
                        End With

258                     MsgBox "Deposit Successfully Added", vbInformation, _
                                "Webplus Lending Corporation"
260                     Unload Me
262                     frm_Deposit.lblUser.Caption = MDIForm1.lblUserName.Caption
264                     frm_Deposit.txtUser.Text = MDIForm1.lblUserName.Caption
266                     Me.Show
                    End If
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnAdd_Click"

        Exit Sub

btnAdd_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.btnAdd_Click", Erl

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
104     ElseIf btnClose.Caption = "&Cancel" Then
106         btnClose.Caption = "&Close"
108         btnEdit.Caption = "&Edit"
110         btnAdd.Caption = "&Add"
112         txtSearch.Enabled = True
114         cbAccount.Text = ""
116         txtTotal.Text = "0"
118         txtAmount.Text = "0"
120         txtParticular.Text = ""
122         btnEdit.Enabled = False
124         btnDelete.Enabled = False
126         dtDate.Enabled = False
128         txtAmount.Enabled = False
130         txtParticular.Enabled = False
132         cbAccount.Enabled = False
134         btnAdd.Enabled = True
136         DataGrid1.Enabled = True
138         txtID.Text = lblAuDep.Caption

140         If rsDeposit.State = 1 Then rsDeposit.Close
142         rsDeposit.Open "Select * from tblDeposit Order by ID desc"
144         Set DataGrid1.DataSource = rsDeposit
146         Call Timer1_Timer
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnDelete_Click()

        '<EhHeader>
        On Error GoTo btnDelete_Click_Err

        TxtLog "Entered btnDelete_Click"

        '</EhHeader>

100     If rsDeposit.State = 1 Then rsDeposit.Close
102     rsDeposit.Open "Select * from tblDeposit where DepositID = '" & txtID.Text & "'"

104     If MsgBox("Are you sure you want to delete this record? ", vbQuestion + _
                vbYesNo, "Webplus Lending Corporation") = vbYes Then
106         rsDeposit.Delete
108         rsDeposit.Update
110         MsgBox "Record has Successfully Deleted.", , "Webplus Lending Corporation"
112         Unload Me
114         Me.Show
         
        End If

        '<EhFooter>

        TxtLog "Exited btnDelete_Click"

        Exit Sub

btnDelete_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.btnDelete_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEdit_Click()

        '<EhHeader>
        On Error GoTo btnEdit_Click_Err

        TxtLog "Entered btnEdit_Click"

        '</EhHeader>

100     If btnEdit.Caption = "&Edit" Then
102         DataGrid1.Enabled = False
104         btnEdit.Caption = "&Update"
106         btnClose.Caption = "&Cancel"
108         dtDate.Enabled = True
110         DataGrid1.Enabled = False
112         cbAccount.Enabled = True
114         txtAmount.Enabled = True
116         txtParticular.Enabled = True
        Else

118         If MsgBox("Are you sure you want to update this record?", vbQuestion + _
                    vbYesNo, "J Lending Corporation") = vbYes Then

120             If rsDeposit.State = 1 Then rsDeposit.Close
122             rsDeposit.Open "Select * from tblDeposit where DepositID = '" & _
                        txtID.Text & "'"

124             With rsDeposit
126                 !DepositID = txtID.Text
128                 !Date = dtDate.Value
130                 !Amount = Format$(Val(txtAmount.Text), "###,###,##0.00")
132                 !AccountName = cbAccount.Text
134                 !Particular = txtParticular.Text
136                 !DateEncoded = txtDateEncoded.Text
138                 !total = Format$(Val(txtTotal.Text), "###,###,##0.00")
140                 !User = txtUser.Text
142                 .Update
                End With
                                
144             If rsTrail.State = 1 Then rsTrail.Close
146             rsTrail.Open "Select * from tblTrail where Username = '" & txtUser.Text _
                        & "'"

148             With rsTrail
150                 .AddNew
152                 !UserName = txtUser.Text
154                 !userlevel = lblUserlevel.Caption
156                 !Activity = "Add New Income Record"
158                 !Time = lblTime.Caption
160                 !Date = dtDate.Value
162                 .Update
                End With

164             Unload Me
166             Call Deposit
168             Me.Show
                        
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnEdit_Click"

        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.btnEdit_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cbAccount_Click()

    '<EhHeader>
    On Error GoTo cbAccount_Click_Err

    TxtLog "Entered cbAccount_Click"

    '</EhHeader>

    '<EhFooter>

    TxtLog "Exited cbAccount_Click"

    Exit Sub

cbAccount_Click_Err:
    ErrReport Err.Description, "LendingClient.frm_Deposit.cbAccount_Click", Erl

    Resume Next

    '</EhFooter>

End Sub

Private Sub cbAccount_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo cbAccount_KeyPress_Err

        TxtLog "Entered cbAccount_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         txtAmount.SetFocus
        Else
104         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited cbAccount_KeyPress"

        Exit Sub

cbAccount_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.cbAccount_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        TxtLog "Entered DataGrid1_Click"

        '</EhHeader>

100     If rsDeposit.RecordCount = 0 Then
        Else

102         With rsDeposit
104             txtID.Text = !DepositID
106             lblID.Caption = !ID
108             cbAccount.Text = !AccountName
110             dtDate.Value = !Date
112             txtParticular.Text = !Particular
114             txtAmount.Text = !Amount
116             txtTotal.Text = txtAmount.Text
            End With

118         btnEdit.Enabled = True
120         btnDelete.Enabled = True
122         btnAdd.Enabled = False
124         btnClose.Caption = "&Cancel"
                
        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_Click"

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.DataGrid1_Click", Erl

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
        ErrReport Err.Description, "LendingClient.frm_Deposit.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call Chart
104     Call Deposit
106     Call Cashonhand
108     Call Loan
   
110     If rsDeposit.State = 1 Then rsDeposit.Close
112     rsDeposit.Open "Select * from tblDeposit Order by ID desc"
114     Call autoNumber
116     Set DataGrid1.DataSource = rsDeposit
118     DataGrid1.Width = Me.Width

120     Call addItems

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    dtDate.Value = Date
    lblTime.Caption = Time
    txtDateEncoded.Text = Date
    
    Timer1.Enabled = False
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtAmount_KeyPress_Err

        TxtLog "Entered txtAmount_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321."

        'Check for Only One Decimal Point In Textbox
102     If KeyAscii = 46 Then

            'If more than one decimal point is typed, only one decimal will be printed
104         If InStr(1, txtAmount.Text, ".") > 0 Then
106             KeyAscii = 0
108             MsgBox "Multiple decimal points are not allowed.", vbOKOnly + _
                        vbInformation

                Exit Sub

            End If
        End If

110     If InStr(1, strvalid, Chr$(KeyAscii)) Then
112     ElseIf KeyAscii = 8 Then
114         KeyAscii = 8
116     ElseIf KeyAscii = 13 Then
    
118         Call btnAdd_Click
        Else
120         KeyAscii = 0
        
        End If

        '<EhFooter>

        TxtLog "Exited txtAmount_KeyPress"

        Exit Sub

txtAmount_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.txtAmount_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsDeposit.State = 1 Then rsDeposit.Close
102     rsDeposit.Open "Select * from tblDeposit where AccountName like '" & _
                txtSearch.Text & "%' or  Particular like '" & txtSearch.Text & _
                "%' or Date like '" & txtSearch.Text & "%' or DepositID like '" & _
                txtSearch.Text & "%' Order By ID desc"
104     Set DataGrid1.DataSource = rsDeposit

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Deposit.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

