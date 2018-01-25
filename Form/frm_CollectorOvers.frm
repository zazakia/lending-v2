VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_CollectorOvers 
   Caption         =   "Collector Over"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Collectors Over"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   20175
      Begin VB.TextBox txtSearchOver 
         Height          =   375
         Left            =   1560
         TabIndex        =   37
         Top             =   4320
         Width           =   6495
      End
      Begin VB.TextBox txtCurrentOR 
         Height          =   285
         Left            =   11280
         TabIndex        =   36
         Text            =   "txtCurrentOR"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtOR 
         Height          =   285
         Left            =   11280
         TabIndex        =   35
         Text            =   "txtOR"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cmbCollCode 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton cmdDeleteOver 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   480
         Left            =   9600
         TabIndex        =   33
         Top             =   2880
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdCancelOver 
         Caption         =   "&Cancel"
         Height          =   480
         Left            =   8040
         TabIndex        =   32
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveOver 
         Caption         =   "&Save"
         Height          =   480
         Left            =   6480
         TabIndex        =   31
         Top             =   2880
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DPOver 
         Height          =   375
         Left            =   7560
         TabIndex        =   30
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   265945089
         CurrentDate     =   42037
      End
      Begin VB.TextBox txtOver 
         Height          =   375
         Left            =   7560
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   7560
         TabIndex        =   22
         Top             =   600
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   360
         Top             =   480
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5295
         Left            =   120
         TabIndex        =   19
         Top             =   4800
         Width           =   19935
         _ExtentX        =   35163
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin VB.TextBox txtAddress 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox txtMI 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   9
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtFirstname 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   2160
         Width           =   3495
      End
      Begin VB.TextBox txtLastname 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label txtDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "txtDate"
         Height          =   195
         Left            =   11280
         TabIndex        =   34
         Top             =   1080
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
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
         Left            =   6840
         TabIndex        =   29
         Top             =   2280
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Over :"
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
         Left            =   6840
         TabIndex        =   28
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label gg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select collector code from dropdown"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Top             =   600
         Width           =   3540
      End
      Begin VB.Label lblLabel17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "search by collector code or date"
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
         Left            =   1560
         TabIndex        =   26
         Top             =   3840
         Width           =   2880
      End
      Begin VB.Label lblAuCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   6120
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label labelhahahahaha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User :"
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
         Index           =   1
         Left            =   6840
         TabIndex        =   23
         Top             =   600
         Width           =   630
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   20040
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Search    :"
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
         TabIndex        =   15
         Top             =   4320
         Width           =   1065
      End
      Begin VB.Label Label14 
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
         Height          =   255
         Left            =   6600
         TabIndex        =   14
         Top             =   6120
         Width           =   135
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
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
         Left            =   9240
         TabIndex        =   13
         Top             =   6720
         Width           =   90
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   255
         Left            =   5160
         TabIndex        =   12
         Top             =   3360
         Width           =   15
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Address        :"
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
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "M.I.                 :"
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
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "First Name   :"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Last Name   :"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code             :"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1125
      End
   End
   Begin VB.Label lblUserr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   10440
      TabIndex        =   24
      Top             =   4080
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblTime 
      Height          =   375
      Left            =   20160
      TabIndex        =   21
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label lblUserlevel 
      AutoSize        =   -1  'True
      Caption         =   "Label16"
      Height          =   195
      Left            =   20160
      TabIndex        =   20
      Top             =   2880
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblLastname 
      AutoSize        =   -1  'True
      Caption         =   "Label16"
      Height          =   195
      Left            =   20280
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCode 
      Caption         =   "Code"
      Height          =   255
      Left            =   20160
      TabIndex        =   17
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9840
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label lblUser 
      Caption         =   "Label16"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   16
      Top             =   7440
      Width           =   1335
   End
End
Attribute VB_Name = "frm_CollectorOvers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Collectors
'    Project    : Project1
'
'    Description: [This procedure will add a New Collector Record]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Sub autoNumber()

        '<EhHeader>
        On Error GoTo autoNumber_Err

        TxtLog "Entered autoNumber"

        '</EhHeader>
               
102     Call payment1

104     If rsPayment1.State = 1 Then rsPayment1.Close
106     rsPayment1.Open "Select Max(ID) from tblPayment"
        'Benjamin Sumilhig And Jun rey Tavera 3-13-2015
        'Di man gyud na mazero nang record Count bisan zero ang value sa rsPayment kay naa man gyud na usa ka record count
        'mao dle modisplay ang OR1
        'if rsPayment.RecordCount = 0 then
        'textOR.Text = "OR1"
108     txtOR.Text = "OR" & rsPayment1(0)

110     If txtOR.Text = "OR" Then
112         txtOR.Text = "OR1"
        Else

            'Dim newOR As Double

            'newOR = rsPayment1(0) + 1
114         txtOR.Text = "OR" & rsPayment1(0) + 1

        End If

        '<EhFooter>

        TxtLog "Exited autoNumber"

        Exit Sub

autoNumber_Err:
        ErrReport Err.Description, "LendingClient.frm_CollectorOvers.autoNumber", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub clear_data()

        '<EhHeader>
        On Error GoTo clear_data_Err

        TxtLog "Entered clear_data"

        '</EhHeader>

100     txtLastname.Text = ""
102     txtFirstname.Text = ""
104     txtMI.Text = ""
106     txtAddress.Text = ""
108     txtOver.Text = ""

        '<EhFooter>

        TxtLog "Exited clear_data"

        Exit Sub

clear_data_Err:
        ErrReport Err.Description, "LendingClient.frm_CollectorOvers.clear_data", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub set_grid()

        '<EhHeader>
        On Error GoTo set_grid_Err

        TxtLog "Entered set_grid"

        '</EhHeader>

100     txtSearchOver.Text = ""

102     If rsPayment.State = 1 Then rsPayment.Close
104     rsPayment.Open _
                "Select top 100 CollectorCode, PaymentsMade, Date, ORnumber from tblPayment where Code = 'Over' and Customer = 'Over' and Status = 'Good' order by ORnumber desc"
        'rsPayment.Open "Select * from tblPayment"
     
106     Set DataGrid1.DataSource = rsPayment
108     DataGrid1.Columns(1).Caption = "Over Collection"
110     DataGrid1.Width = Me.Width
112     DataGrid1.Refresh

        '<EhFooter>

        TxtLog "Exited set_grid"

        Exit Sub

set_grid_Err:
        ErrReport Err.Description, "LendingClient.frm_CollectorOvers.set_grid", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmbCollCode_Click()

        '<EhHeader>
        On Error GoTo cmbCollCode_Click_Err

        TxtLog "Entered cmbCollCode_Click"

        '</EhHeader>

        'ani i open na ang collector data
100     If rsCollData.State = 1 Then rsCollData.Close
102     rsCollData.Open "Select * from tblColl_Data where Code = '" & cmbCollCode.Text _
                & "' ORDER BY DateEmployed Desc"

104     If rsCollData.RecordCount <> 0 Then
106         txtLastname.Text = rsCollData!lastname
108         txtFirstname.Text = rsCollData!firstname
110         txtMI.Text = rsCollData!MI
112         txtAddress.Text = rsCollData!Adress
        End If

        '<EhFooter>

        TxtLog "Exited cmbCollCode_Click"

        Exit Sub

cmbCollCode_Click_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.cmbCollCode_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmbCollCode_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo cmbCollCode_KeyPress_Err

        TxtLog "Entered cmbCollCode_KeyPress"

        '</EhHeader>

100     If KeyAscii = 40 Then 'KeyAscii = 38 Or
102         If rsCollData.State = 1 Then rsCollData.Close
104         rsCollData.Open "Select * from tblColl_Data "
106         cmbCollCode.Clear

108         If rsCollData.RecordCount <> 0 Then

110             Do While Not rsCollData.EOF
112                 cmbCollCode.AddItem rsCollData!lastname
114                 rsCollData.MoveNext
                Loop

            End If
        End If

        '<EhFooter>

        TxtLog "Exited cmbCollCode_KeyPress"

        Exit Sub

cmbCollCode_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.cmbCollCode_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdCancelOver_Click()

        '<EhHeader>
        On Error GoTo cmdCancelOver_Click_Err

        TxtLog "Entered cmdCancelOver_Click"

        '</EhHeader>

100     Call set_grid
102     Call clear_data
104     txtCurrentOR.Text = ""
106     cmbCollCode.Enabled = True
108     DPOver.Enabled = True
110     cmdSaveOver.Caption = "&Save"
112     cmdDeleteOver.Visible = False
114     cmdDeleteOver.Enabled = False

        '<EhFooter>

        TxtLog "Exited cmdCancelOver_Click"

        Exit Sub

cmdCancelOver_Click_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.cmdCancelOver_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdDeleteOver_Click()

        '<EhHeader>
        On Error GoTo cmdDeleteOver_Click_Err

        TxtLog "Entered cmdDeleteOver_Click"

        '</EhHeader>

100     If rsPayment1.State = 1 Then rsPayment1.Close
102     rsPayment1.Open "Select * from tblPayment where ORnumber = '" & _
                txtCurrentOR.Text & "'"

104     If rsPayment1.RecordCount <> 0 Then
106         If MsgBox("Are you want delete this over collection record?", vbQuestion + _
                    vbYesNo, "Webplus Lending Corporation") = vbYes Then
108             rsPayment1.Delete
                'rsPayment1.Update
110             rsPayment1.Update
                'how do we know if a record is successfully deleted?
112             Call cmdCancelOver_Click
            End If

        Else
114         MsgBox ("Record not found")
        End If

        'rsPayment.Update

        '<EhFooter>

        TxtLog "Exited cmdDeleteOver_Click"

        Exit Sub

cmdDeleteOver_Click_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.cmdDeleteOver_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdSaveOver_Click()

        '<EhHeader>
        On Error GoTo cmdSaveOver_Click_Err

        TxtLog "Entered cmdSaveOver_Click"

        '</EhHeader>

100     If txtOver.Text = "" Or Val(txtOver.Text) = 0 Then
102         MsgBox ("Amount should not be zero or empty")
        Else

104         If cmdSaveOver.Caption = "&Save" Then
106             If cmbCollCode.Text = "" Or txtLastname.Text = "" Then
108                 MsgBox ("Please Select a collector first")
                Else
110                 Call autoNumber

112                 If rsPayment1.State = 1 Then rsPayment1.Close
114                 rsPayment1.Open "Select * from tblPayment where CollectorCode = '" _
                            & cmbCollCode.Text & "' and Date = #" & DPOver.Value & _
                            "# and Code = 'Over' and Status = 'Good'"

116                 If rsPayment1.RecordCount = 0 Then
118                     If MsgBox("Are you sure?", vbQuestion + vbYesNo, _
                                "Webplus Lending Corporation") = vbYes Then

120                         With rsPayment1
122                             .AddNew
124                             !ORnumber = txtOR.Text
126                             !LoanID = 0
128                             !Date = DPOver.Value
130                             !DateEncoded = txtDate.Caption
132                             !Collector = txtLastname.Text
134                             !code = "Over"
                                'LCode = lblCuCode.Caption
136                             !Customer = "Over"
138                             !principal = 0
140                             !DateRelease = txtDate.Caption
142                             !Maturity = txtDate.Caption
144                             !Status = "Good"
146                             !Amortization = 0
148                             !paymentsMade = Val(txtOver.Text)
150                             !TotalBalance = 0
152                             !NewBalance = 0
154                             !User = txtUser.Text
                                '!Over = Val(lblOver.Caption)
156                             !TotalPayment = 0
158                             !CollectorCode = cmbCollCode.Text
160                             !CollectorFname = txtFirstname.Text
162                             .Update
                            End With

164                         Call cmdCancelOver_Click
                        
                        End If

                    Else
166                     MsgBox ("Collector " + cmbCollCode.Text + _
                                " already has an over for the chosen date")
                    End If
                End If
           
168         ElseIf cmdSaveOver.Caption = "&Update" Then

170             If MsgBox("Are you sure you want edit the current record?", vbQuestion _
                        + vbYesNo, "Webplus Lending Corporation") = vbYes Then

172                 If rsPayment1.State = 1 Then rsPayment1.Close
174                 rsPayment1.Open "Select * from tblPayment where ORnumber = '" & _
                            txtCurrentOR.Text & "'"

176                 If rsPayment1.RecordCount <> 0 Then
                    
178                     With rsPayment1

180                         If !paymentsMade <> Val(txtOver.Text) Then
182                             !paymentsMade = Val(txtOver.Text)
184                             !User = txtUser.Text
186                             .Update
188                             MsgBox ("Update Complete")
190                             Call cmdCancelOver_Click
                            Else
192                             MsgBox ("Please change the over collection amount")
                            End If

                        End With

                    Else
194                     MsgBox ("Cannot find record to update")
196                     Call cmdCancelOver_Click
                    End If

                    'Call autonumber
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited cmdSaveOver_Click"

        Exit Sub

cmdSaveOver_Click_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.cmdSaveOver_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        TxtLog "Entered DataGrid1_Click"

        '</EhHeader>

        '
100     If rsPayment.RecordCount <> 0 Then
        
102         With rsPayment
104             cmbCollCode.Text = !CollectorCode
106             txtOver.Text = !paymentsMade
108             DPOver.Value = !Date
110             txtCurrentOR.Text = !ORnumber
            End With

            'get collector info using collector code
                
112         If rsCollData.State = 1 Then rsCollData.Close
114         rsCollData.Open "Select * from tblColl_Data where Code = '" & _
                    cmbCollCode.Text & "'"
                
116         If rsCollData.RecordCount <> 0 Then
118             cmbCollCode.Text = rsCollData!code
120             cmdSaveOver.Caption = "&Update"
122             cmbCollCode.Enabled = False
124             DPOver.Enabled = False
126             cmdDeleteOver.Visible = True
128             cmdDeleteOver.Enabled = True
            Else
130             MsgBox ("collector for the over collection not found.")
132             Call clear_data
134             Call set_grid
            End If

            'set button values here'
                
        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_Click"

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_CollectorOvers.DataGrid1_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DataGrid1_KeyPress_Err

        TxtLog "Entered DataGrid1_KeyPress"

        '</EhHeader>

100     KeyAscii = 0 ' para di ma edit ang datagrid?
        ' what about sa copy paste

        '<EhFooter>

        TxtLog "Exited DataGrid1_KeyPress"

        Exit Sub

DataGrid1_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                Y As Single)

        '<EhHeader>
        On Error GoTo DataGrid1_MouseDown_Err

        TxtLog "Entered DataGrid1_MouseDown"

        '</EhHeader>

100     If Button = vbRightButton Then
102         DataGrid1.AllowUpdate = False
        Else
104         DataGrid1.AllowUpdate = True
        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_MouseDown"

        Exit Sub

DataGrid1_MouseDown_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.DataGrid1_MouseDown", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

        'march 13 2015
        'Benjamin Sumilhig
        'this will call the record set of colldata and collcode
100     Call CollData
102     Call CollCode
104     Call connect
106     Call Collector
108     Call Customer
110     Call autoNumber
112     Call payment
114     Call Loan
116     Call payment1
        
118     Call set_grid
                
120     If rsCollCode.State = 1 Then rsCollCode.Close
122     rsCollCode.Open "Select * from tblColl_Code"
124     cmbCollCode.Clear
        
126     If rsCollCode.RecordCount <> 0 Then
                
128         Do While Not rsCollCode.EOF
130             cmbCollCode.AddItem rsCollCode!code
132             rsCollCode.MoveNext
            Loop
                
        End If

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_CollectorOvers.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    txtDate.Caption = Date
    DPOver.Value = Date
    lblTime.Caption = Time
    
    Timer1.Enabled = False
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtAddress_KeyPress_Err

        TxtLog "Entered txtAddress_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then

        End If

        '<EhFooter>

        TxtLog "Exited txtAddress_KeyPress"

        Exit Sub

txtAddress_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.txtAddress_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtFirstname_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtFirstname_KeyPress_Err

        TxtLog "Entered txtFirstname_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz- "

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then

104     ElseIf KeyAscii = 13 Then

106         txtMI.SetFocus
 
108     ElseIf KeyAscii = 8 Then
110         KeyAscii = 8

        Else
112         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtFirstname_KeyPress"

        Exit Sub

txtFirstname_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.txtFirstname_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtMI_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtMI_KeyPress_Err

        TxtLog "Entered txtMI_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then

104     ElseIf KeyAscii = 13 Then

106         txtAddress.SetFocus
 
108     ElseIf KeyAscii = 8 Then
110         KeyAscii = 8

        Else
112         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtMI_KeyPress"

        Exit Sub

txtMI_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_CollectorOvers.txtMI_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtOver_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtOver_KeyPress_Err

        TxtLog "Entered txtOver_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321"

        '        'Check for Only One Decimal Point In Textbox
        '106     If KeyAscii = 46 Then
        '
        '            'If more than one decimal point is typed, only one decimal will be printed
        '108         If InStr(1, txtAmountPaid.Text, ".") > 0 Then
        '110             KeyAscii = 0
        '112             MsgBox "Multiple decimal points are not allowed.", vbOKOnly + vbInformation
        '
        '                Exit Sub
        '
        '            End If

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8

108     ElseIf KeyAscii = 13 Then

            'txtTotaltAmortization.Text = txtAmortization.Text
            'Call Compute
            '   btnAddpayment.SetFocus
            'dtDate.SetFocus
            'Call btnAddpayment_Click
            
        Else
110         KeyAscii = 0
        
        End If

        '<EhFooter>

        TxtLog "Exited txtOver_KeyPress"

        Exit Sub

txtOver_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_CollectorOvers.txtOver_KeyPress", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtSearchOver_Change()

        '<EhHeader>
        On Error GoTo txtSearchOver_Change_Err

        TxtLog "Entered txtSearchOver_Change"

        '</EhHeader>

100     If txtSearchOver.Text <> "" Then
102         If rsPayment.State = 1 Then rsPayment.Close
104         rsPayment.Open _
                    "Select top 100 CollectorCode, PaymentsMade, Date, ORnumber from tblPayment where Code = 'Over' and Customer = 'Over' and Status = 'Good' and ((CollectorCode like '%" _
                    & txtSearchOver.Text & "%') or (Date like '" & txtSearchOver.Text & _
                    "%'))"
            ' or CollectorCode like '" & txtSearchOver.Text & "%' or Date = #" & txtSearchOver.Text & "#"
106         Set DataGrid1.DataSource = rsPayment
108         DataGrid1.Columns(1).Caption = "Over Collection"
110         DataGrid1.Width = Me.Width
112         DataGrid1.Refresh
        Else
114         Call set_grid
        End If

        '<EhFooter>

        TxtLog "Exited txtSearchOver_Change"

        Exit Sub

txtSearchOver_Change_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.txtSearchOver_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtSearchOver_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtSearchOver_KeyPress_Err

        TxtLog "Entered txtSearchOver_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321/"

102     If (InStr(1, strvalid, Chr$(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 13) = 0 Then
104         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtSearchOver_KeyPress"

        Exit Sub

txtSearchOver_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CollectorOvers.txtSearchOver_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

