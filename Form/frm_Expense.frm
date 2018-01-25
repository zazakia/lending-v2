VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Expenses 
   Caption         =   "frmExpense"
   ClientHeight    =   10080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "form1"
   ScaleHeight     =   10080
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm_Expense 
      Caption         =   "Expense"
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
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   29
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   600
         Left            =   5640
         TabIndex        =   27
         Top             =   9240
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtFilter 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   530055169
         CurrentDate     =   41897
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   600
         Left            =   3000
         TabIndex        =   23
         Top             =   9240
         Width           =   2055
      End
      Begin VB.TextBox txtDateEncoded 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   11880
         TabIndex        =   22
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtParticular 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   2880
         Width           =   3015
      End
      Begin VB.ComboBox cbAccount 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   16
         Top             =   2400
         Width           =   3015
      End
      Begin VB.TextBox txtID 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   3960
         Width           =   5175
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   5880
         Top             =   1680
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   530055169
         CurrentDate     =   41873
      End
      Begin VB.TextBox txtAmount 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "0"
         Top             =   1920
         Width           =   3015
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   8160
         TabIndex        =   3
         Top             =   9240
         Width           =   2055
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   9240
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4695
         Left            =   120
         TabIndex        =   1
         Top             =   4440
         Width           =   19815
         _ExtentX        =   34951
         _ExtentY        =   8281
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
      Begin VB.Label lblRR 
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
         Left            =   1920
         TabIndex        =   33
         Top             =   3480
         Width           =   90
      End
      Begin VB.Label lblLabel9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH USING ACCOUNT NAME , PARTICULARS OR DATE"
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
         Left            =   2040
         TabIndex        =   32
         Top             =   3480
         Width           =   6825
      End
      Begin VB.Label lbl 
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
         Left            =   5040
         TabIndex        =   31
         Top             =   2880
         Width           =   90
      End
      Begin VB.Label lblAuEx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   9960
         TabIndex        =   30
         Top             =   1680
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "User: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5880
         TabIndex        =   28
         Top             =   360
         Width           =   525
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   20040
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Date :"
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
         Left            =   960
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   1230
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
         Left            =   11160
         TabIndex        =   21
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblAccountNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account ID :"
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
         TabIndex        =   14
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Search Here      :"
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
         TabIndex        =   11
         Top             =   3960
         Width           =   1770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Date:  "
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
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Amount  :"
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
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Account Name :"
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
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Particulars :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   1335
      End
   End
   Begin VB.Label lblACno 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   10200
      TabIndex        =   26
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   20400
      TabIndex        =   19
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label lblUserlevel 
      Caption         =   "Label6"
      Height          =   255
      Left            =   20280
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblTime 
      Caption         =   "Label6"
      Height          =   255
      Left            =   20280
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   9960
      TabIndex        =   13
      Top             =   4800
      Width           =   45
   End
   Begin VB.Label lblType 
      Height          =   135
      Left            =   16080
      TabIndex        =   12
      Top             =   6480
      Width           =   135
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   4560
   End
End
Attribute VB_Name = "frm_Expenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Expense
'    Project    : Project1
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Sub adddescription()

    '<EhHeader>
    On Error GoTo adddescription_Err

    TxtLog "Entered adddescription"

    '</EhHeader>

    '  cbParticulars.Clear

    '   If rsChart.RecordCount <> 0 Then

    '        Do While Not rsChart.EOF
    '         cbParticulars.AddItem rsChart!Description
    '         rsChart.MoveNext
    '       Loop

    '   End If

    '<EhFooter>

    TxtLog "Exited adddescription"

    Exit Sub

adddescription_Err:
    ErrReport Err.Description, "LendingClient.frm_Expenses.adddescription", Erl

    Resume Next

    '</EhFooter>

End Sub

Sub addItems()

        '<EhHeader>
        On Error GoTo addItems_Err

        TxtLog "Entered addItems"

        '</EhHeader>

100     If rsChart.State = 1 Then rsChart.Close
102     rsChart.Open "Select * from tblChartOfAccounts  where Type = '" & "Expense" & "'"
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
        ErrReport Err.Description, "LendingClient.frm_Expenses.addItems", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub autoNumber()

        '<EhHeader>
        On Error GoTo autoNumber_Err

        TxtLog "Entered autoNumber"

        '</EhHeader>

100     If rsExpense.RecordCount = 0 Then
102         txtID.Text = "E01"
        Else
            '106         rsExpense.MoveLast
            '108         txtID.Text = "E" & Format(Right(rsExpense!AccountID, 3) + 1, "000")
104         txtID.Text = "E" & rsExpense!ID
106         lblAuEx.Caption = txtID.Text
        End If

        '<EhFooter>

        TxtLog "Exited autoNumber"

        Exit Sub

autoNumber_Err:
        ErrReport Err.Description, "LendingClient.frm_Expenses.autoNumber", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnAdd_Click()

        '<EhHeader>
        On Error GoTo btnAdd_Click_Err

        TxtLog "Entered btnAdd_Click"

        '</EhHeader>

100     If btnAdd.Caption = "&Add" Then
102         btnAdd.Caption = "&Save"
104         btnClose.Caption = "&Cancel"
106         txtAmount.Enabled = True
108         cbAccount.Enabled = True
110         txtParticular.Enabled = True
112         dtDate.Enabled = True
114     ElseIf btnAdd.Caption = "&Save" Then

116         If rsExpense.State = 1 Then rsExpense.Close
118         rsExpense.Open "Select * from tblExpense where AccountID = '" & _
                    cbAccount.Text & "'"

120         If rsExpense.RecordCount <> 0 Then
122         ElseIf txtParticular.Text = "" Then
124             MsgBox "Particular Must not be blank", vbInformation, _
                        "Webplus Lending Corporation"
126         ElseIf cbAccount.Text = "" Then
128             MsgBox "Account Name Must not be blank", vbInformation, _
                        "Webplus Lending Corporation"
            Else

130             If MsgBox("Are you sure you Add this Expense?", vbQuestion + vbYesNo) = _
                        vbYes Then

                    'Benjamin Sumilhig
132                 If rsLoan.State = 1 Then rsLoan.Close
134                 rsLoan.Open "SELECT * FROM tblLoan WHERE DateRelease = #" & Format$( _
                            dtDate.Value, "mm/dd/yy") & "# And Status = 'Good'"

136                 If rsLoan.RecordCount = 0 Then

                        'MsgBox "=0"
138                     With rsLoan
140                         .AddNew
142                         !Collector = " "
144                         !code = " "
146                         !Customer = " "
148                         !firstname = " "
150                         !principal = 0
152                         !total = 0
154                         !LoanDate = Format$(dtDate.Value, "m/d/yy")
156                         !DateRelease = Format$(dtDate.Value, "m/d/yy")
158                         !Maturity = Format$(dtDate.Value, "m/d/yy")
160                         !Status = "Good"
162                         !FireInsurance = 0
164                         !CollectorCharge = 0
166                         !delivery = 0
168                         !collection = 0
170                         !Servicefee = 0
172                         !Balance = 0
174                         !Penalty = 0
176                         !Passbook = 0
178                         !TotalAmortization = 0
180                         !TotalCharges = 0
182                         !LoanTotal = 0
184                         !CollectorCode = " "
186                         !CollectorFname = " "
188                         !User = " "
190                         !LoanStatus = "Good"
192                         !TotalPayment = 0
194                         !NotPosted = 0
196                         !OverToCus = 0
198                         .Update
                        End With

                    Else
                        'Blanko lang ning else
                        'MsgBox ">0"
                    End If

                    'End of Benjamin Sumilhig

200                 With rsExpense
202                     .AddNew
204                     !AccountID = txtID.Text
206                     !ToDate = dtDate.Value
208                     !AccountName = Trim$(cbAccount.Text)
210                     !Amount = Format$(Val(txtAmount.Text), "###,###,##0.00")
212                     !ToDate = dtDate.Value
214                     !Particular = Trim$(txtParticular.Text)
216                     !ChartID = lblACno.Caption
218                     !User = Trim$(txtUser.Text)
220                     !DateEncoded = CDate(txtDateEncoded.Text)
222                     .Update
                    End With
                
224                 If rsTrail.State = 1 Then rsTrail.Close
226                 rsTrail.Open "Select * from tblTrail where Username = '" & _
                            lblUser.Caption & "'"

228                 With rsTrail
230                     .AddNew
232                     !UserName = txtUser.Text
234                     !userlevel = lblUserlevel.Caption
236                     !Activity = "Add New Expense Record"
238                     !Time = lblTime.Caption
240                     !Date = dtDate.Value
242                     .Update
                    End With

244                 MsgBox "Expense successfully Added", vbInformation
246                 Unload Me
248                 frm_Expenses.lblUser.Caption = MDIForm1.lblUserName.Caption
250                 frm_Expenses.txtUser.Text = MDIForm1.lblUserName.Caption
252                 Me.Show
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnAdd_Click"

        Exit Sub

btnAdd_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Expenses.btnAdd_Click", Erl

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
108         btnAdd.Caption = "&Add"
110         btnEdit.Caption = "&Edit"
112         btnEdit.Enabled = False
114         btnAdd.Enabled = True
116         btnDelete.Enabled = False
118         txtSearch.Enabled = True
120         cbAccount.Text = ""
            
            ' txtTotal.Text = "0"
122         dtDate.Enabled = False
124         txtAmount.Text = "0"
126         txtParticular.Text = ""
128         cbAccount.Enabled = False
130         txtAmount.Enabled = False
132         txtParticular.Enabled = False
134         txtID.Text = lblAuEx.Caption
136         DataGrid1.Enabled = True
138         Call Timer1_Timer
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Expenses.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnDelete_Click()

        '<EhHeader>
        On Error GoTo btnDelete_Click_Err

        TxtLog "Entered btnDelete_Click"

        '</EhHeader>

100     If rsExpense.State = 1 Then rsExpense.Close
102     rsExpense.Open "Select * from tblExpense where AccountID = '" & txtID.Text & "'"

        '  If Val(rsExpense!Balance) > 0 Then
        '         MsgBox "This customer has a remaining balance. You can't delete it.", vbInformation, "Webplus Lending Corporation"
        '        Call Expense
        '       Set DataGrid1.DataSource = rsExpense
        '  Else

104     If MsgBox("Are you sure you want to delete this record? ", vbQuestion + _
                vbYesNo, "Webplus Lending Corporation") = vbYes Then
106         rsExpense.Delete
108         rsExpense.Update
110         MsgBox "Record has Successfully Deleted.", , "Webplus Lending Corporation"
112         Unload Me
114         Me.Show
            '  Load frm_Expense
            '  frm_Expense.Show
        End If

        '    End If

        '<EhFooter>

        TxtLog "Exited btnDelete_Click"

        Exit Sub

btnDelete_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Expenses.btnDelete_Click", Erl

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
104         btnClose.Caption = "&Cancel"
106         dtDate.Enabled = True
108         txtAmount.Enabled = True
110         cbAccount.Enabled = True
112         txtParticular.Enabled = True
114         DataGrid1.Enabled = False
        Else

116         If MsgBox("Are you sure you want to update this record?", vbQuestion + _
                    vbYesNo, "J Lending Corporation") = vbYes Then

118             If rsExpense.State = 1 Then rsExpense.Close
120             rsExpense.Open "Select * from tblExpense where AccountID = '" & _
                        txtID.Text & "'"

122             With rsExpense
             
124                 !AccountID = txtID.Text
126                 !ToDate = dtDate.Value
128                 !AccountName = cbAccount.Text
130                 !Amount = Format$(Val(txtAmount.Text), "###,###,##0.00")
132                 !Particular = txtParticular.Text
134                 !DateEncoded = txtDateEncoded.Text
136                 !User = Trim$(txtUser.Text)
138                 .Update
                End With

140             If rsTrail.State = 1 Then rsTrail.Close
142             rsTrail.Open "Select * from tblTrail where Username = '" & _
                        lblUser.Caption & "'"

144             With rsTrail
146                 .AddNew
148                 !UserName = txtUser.Text
150                 !userlevel = lblUserlevel.Caption
152                 !Activity = "Add New Expense Record"
154                 !Time = lblTime.Caption
156                 !Date = dtDate.Value
158                 .Update
                End With
                
160             Unload Me
162             Call Expense
164             Me.Show
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnEdit_Click"

        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Expenses.btnEdit_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cbAccount_Click()

        '<EhHeader>
        On Error GoTo cbAccount_Click_Err

        TxtLog "Entered cbAccount_Click"

        '</EhHeader>

100     If rsChart.State = 1 Then rsChart.Close
102     rsChart.Open "Select * from tblChartOfAccounts where AccountName = '" & _
                cbAccount.Text & "'"
104     lblACno.Caption = rsChart!AccountNo
106     Call adddescription

        '<EhFooter>

        TxtLog "Exited cbAccount_Click"

        Exit Sub

cbAccount_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Expenses.cbAccount_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cbAccount_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo cbAccount_KeyPress_Err

        TxtLog "Entered cbAccount_KeyPress"

        '</EhHeader>

100     KeyAscii = 0

        '<EhFooter>

        TxtLog "Exited cbAccount_KeyPress"

        Exit Sub

cbAccount_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Expenses.cbAccount_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        TxtLog "Entered DataGrid1_Click"

        '</EhHeader>

100     If rsExpense.RecordCount = 0 Then
        Else

102         With rsExpense
104             dtDate.Value = !ToDate
106             cbAccount.Text = !AccountName
108             txtAmount.Text = !Amount
110             txtID.Text = !AccountID
                
112             txtParticular.Text = !Particular
        
            End With

114         btnEdit.Enabled = True
116         btnAdd.Enabled = False
118         btnDelete.Enabled = True
120         btnClose.Caption = "&Cancel"
        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_Click"

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Expenses.DataGrid1_Click", Erl

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
        ErrReport Err.Description, "LendingClient.frm_Expenses.DataGrid1_KeyPress", Erl

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
104     Call Expense
106     Call Loan
        
108     If rsExpense.State = 1 Then rsExpense.Close
110     rsExpense.Open "Select * from tblExpense Order By ID desc"
112     Call autoNumber
114     Set DataGrid1.DataSource = rsExpense
116     DataGrid1.Width = Me.Width
118     DataGrid1.Columns(0).Width = 400
120     DataGrid1.Columns(1).Width = 900
        
122     Call addItems

        'Call adddescription

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Expenses.Form_Load", Erl

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
        ErrReport Err.Description, "LendingClient.frm_Expenses.txtAmount_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsExpense.State = 1 Then rsExpense.Close
102     rsExpense.Open "Select * from tblExpense where AccountName like '" & _
                txtSearch.Text & "%' or Particular like '" & txtSearch.Text & _
                "%' or ToDate like '" & txtSearch.Text & "%' or AccountID like '" & _
                txtSearch.Text & "%' Order by ID desc"
104     Set DataGrid1.DataSource = rsExpense

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Expenses.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

