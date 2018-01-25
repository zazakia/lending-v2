VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_CashOnHand 
   Caption         =   "frm Cash On Hand & Cash on Bank"
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20025
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
   Moveable        =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   20025
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      Format          =   96600065
      CurrentDate     =   41890
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8040
      Top             =   480
   End
   Begin VB.Frame Frame1 
      Height          =   10095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   19695
      Begin VB.Frame Frame5 
         Height          =   975
         Left            =   7200
         TabIndex        =   27
         Top             =   9000
         Width           =   2055
         Begin VB.CommandButton btnClose1 
            Caption         =   "&Close"
            Height          =   735
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnEditCashB 
         Caption         =   "&Edit Cash On Bank"
         Enabled         =   0   'False
         Height          =   720
         Left            =   12000
         TabIndex        =   19
         Top             =   9120
         Width           =   1815
      End
      Begin VB.CommandButton cmdEditCash 
         Caption         =   "&Edit Cash on Hand"
         Enabled         =   0   'False
         Height          =   720
         Left            =   2160
         TabIndex        =   18
         Top             =   9120
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtOnBank 
         Height          =   375
         Left            =   17040
         TabIndex        =   15
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   96600065
         CurrentDate     =   41897
      End
      Begin MSComCtl2.DTPicker dtOnHand 
         Height          =   375
         Left            =   6960
         TabIndex        =   14
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   96600065
         CurrentDate     =   41897
      End
      Begin VB.CommandButton btnEditBank 
         Caption         =   "&Add Cash on Bank"
         Height          =   720
         Left            =   9960
         TabIndex        =   13
         Top             =   9120
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   7215
         Left            =   9840
         TabIndex        =   12
         Top             =   1440
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   12726
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.TextBox txtCashonbank 
         Enabled         =   0   'False
         Height          =   375
         Left            =   12120
         TabIndex        =   10
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   840
         Left            =   17760
         TabIndex        =   5
         Top             =   9000
         Width           =   1710
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Add Cash on Hand"
         Height          =   720
         Left            =   240
         TabIndex        =   4
         Top             =   9120
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   12726
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
               LCID            =   13321
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
               LCID            =   13321
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
      Begin VB.TextBox txtCashOnHand 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   960
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   8880
         Width           =   3855
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   9840
         TabIndex        =   21
         Top             =   8880
         Width           =   4095
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   17640
         TabIndex        =   22
         Top             =   8880
         Width           =   1935
      End
      Begin VB.Label lbltytyrty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   14760
         TabIndex        =   33
         Top             =   960
         Width           =   135
      End
      Begin VB.Label ererer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   5040
         TabIndex        =   32
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   30
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Must be typed by numbers only."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   29
         Top             =   240
         Width           =   3870
      End
      Begin VB.Label lblTransactionDate2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   15480
         TabIndex        =   17
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label lblTransactionDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5400
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash On Bank  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         TabIndex        =   11
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label rrr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9120
         TabIndex        =   8
         Top             =   240
         Width           =   570
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hhhhhhhh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9960
         TabIndex        =   7
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblCashOn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash On Hand :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   31
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label lblIDb 
      Caption         =   "Label3"
      Height          =   255
      Left            =   18360
      TabIndex        =   26
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblID 
      Caption         =   "lblID"
      Height          =   135
      Left            =   17760
      TabIndex        =   25
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblOnBank 
      Caption         =   "Label3"
      Height          =   255
      Left            =   18000
      TabIndex        =   24
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblOnHand 
      Caption         =   "Label2"
      Height          =   255
      Left            =   18120
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      Height          =   195
      Left            =   7440
      TabIndex        =   9
      Top             =   960
      Width           =   330
   End
End
Attribute VB_Name = "frm_CashOnHand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose1_Click()

        '<EhHeader>
        On Error GoTo btnClose1_Click_Err

        TxtLog "Entered btnClose1_Click"

        '</EhHeader>

100     If btnClose1.Caption = "&Close" Then
102         Unload Me
104     ElseIf btnClose1.Caption = "&Cancel" Then
106         btnClose1.Caption = "&Close"
108         btnEdit.Caption = "&Add Cash on Hand"
110         btnEditBank.Caption = "&Add Cash on Bank"
112         DataGrid1.Enabled = True
114         cmdEditCash.Caption = "&Edit Cash on Hand"
116         btnEditCashB.Enabled = False
118         cmdEditCash.Enabled = False
120         txtCashOnHand.Text = ""
122         txtCashonbank.Text = ""
124         btnEdit.Enabled = True
126         dtOnHand.Enabled = False
128         btnEditBank.Enabled = True
130         DataGrid2.Enabled = True
132         txtCashOnHand.Enabled = False
134         txtCashonbank.Enabled = False
136         dtOnHand.Value = Date
        End If

        '<EhFooter>

        TxtLog "Exited btnClose1_Click"

        Exit Sub

btnClose1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.btnClose1_Click", Erl

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
108         btnEdit.Caption = "&Add Cash on Hand"
110         btnEditBank.Caption = "&Add Cash on Bank"
112         btnEditCashB.Caption = "&Edit Cash On Bank"
114         DataGrid2.Enabled = True
116         DataGrid1.Enabled = True
118         btnEditCashB.Enabled = False
120         cmdEditCash.Enabled = False
122         txtCashOnHand.Text = ""
124         txtCashonbank.Text = ""
126         btnEdit.Enabled = True
128         dtOnHand.Enabled = False
130         btnEditBank.Enabled = True
132         DataGrid2.Enabled = True
134         txtCashOnHand.Enabled = False
136         txtCashonbank.Enabled = False
138         dtOnBank.Value = Date
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEdit_Click()

        '<EhHeader>
        On Error GoTo btnEdit_Click_Err

        TxtLog "Entered btnEdit_Click"

        '</EhHeader>

100     If btnEdit.Caption = "&Add Cash on Hand" Then
            
102         btnEdit.Caption = "&Save"
104         btnClose.Caption = "&Cancel"
106         btnClose1.Caption = "&Cancel"
108         txtCashOnHand.Enabled = True
110         txtCashOnHand.SetFocus
112         dtOnHand.Enabled = True
114         DataGrid1.Enabled = False
116     ElseIf btnEdit.Caption = "&Save" Then

118         If txtCashOnHand.Text = "" Then
120             MsgBox "Field must not be blank.", vbInformation, _
                        "Webplus Lending Corporation"
122             txtCashOnHand.SetFocus

124             If rsCashOnHand.State = 1 Then rsCashOnHand.Close
126             rsCashOnHand.Open "Select * from tblCashOnHand Order by ID desc"
128             Set DataGrid1.DataSource = rsCashOnHand
            Else

130             If rsCashOnHand.State = 1 Then rsCashOnHand.Close
132             rsCashOnHand.Open _
                        "Select * from tblCashOnHand where Transactiondate = #" & _
                        dtOnHand.Value & "#"
                    
134             If rsCashOnHand.RecordCount <> 0 Then
136                 MsgBox _
                            "Cash on hand is already encoded on this date.Just edit if you wish to change the value.", _
                            vbInformation
138                 txtCashOnHand.SetFocus

140                 If rsCashOnHand.State = 1 Then rsCashOnHand.Close
142                 rsCashOnHand.Open "Select * from tblCashOnHand Order by ID desc"
144                 Set DataGrid1.DataSource = rsCashOnHand
                     
                Else

146                 If MsgBox("Are you sure you want to update cash on hand balance?", _
                            vbQuestion + vbYesNo) = vbYes Then

148                     With rsCashOnHand
150                         .AddNew
152                         !Cashonhand = Format$(Val(txtCashOnHand.Text), _
                                    "###,###,##0.00")
154                         !UpdateDate = lblDate.Caption
156                         !User = lblUser.Caption
158                         !TransactionDate = dtOnHand.Value
160                         .Update
                        End With
                     
162                     Unload Me
164                     frm_CashOnHand.lblUser.Caption = MDIForm1.lblUserName.Caption
166                     Me.Show
                    Else

168                     If rsCashOnHand.State = 1 Then rsCashOnHand.Close
170                     rsCashOnHand.Open "Select * from tblCashOnHand Order by ID desc"
172                     Set DataGrid1.DataSource = rsCashOnHand
                    End If
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnEdit_Click"

        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.btnEdit_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEditBank_Click()

        '<EhHeader>
        On Error GoTo btnEditBank_Click_Err

        TxtLog "Entered btnEditBank_Click"

        '</EhHeader>

100     If btnEditBank.Caption = "&Add Cash on Bank" Then
102         btnEditBank.Caption = "&Save"
            ' btnClose.Caption = "&Cancel"
104         btnEditBank.Enabled = True
106         txtCashonbank.Enabled = True
108         txtCashonbank.SetFocus
110         btnClose.Caption = "&Cancel"
112         DataGrid2.Enabled = False
114     ElseIf btnEditBank.Caption = "&Save" Then

116         If txtCashonbank.Text = "" Then
118             MsgBox "Field must not be blank.", vbInformation, _
                        "Webplus Lending Corporation"
120             txtCashonbank.SetFocus

122             If rsCashOnBank.State = 1 Then rsCashOnBank.Close
124             rsCashOnBank.Open "Select * from tblCashOnBank Order by ID desc"
126             Set DataGrid2.DataSource = rsCashOnBank
            Else

128             If MsgBox("Are you sure you want to update cash on Bank balance?", _
                        vbQuestion + vbYesNo) = vbYes Then

130                 If rsCashOnBank.State = 1 Then rsCashOnBank.Close
132                 rsCashOnBank.Open _
                            "Select * from tblCashOnBank where TransactionDate = #" & _
                            dtOnBank.Value & "#"

134                 If rsCashOnBank.RecordCount <> 0 Then
136                     MsgBox _
                                "Cash on Bank is already encoded on this date.Just edit if you wish to change the value.", _
                                vbInformation

138                     If rsCashOnBank.State = 1 Then rsCashOnBank.Close
140                     rsCashOnBank.Open "Select * from tblCashOnBank Order by ID desc"
142                     Set DataGrid2.DataSource = rsCashOnBank
                    Else

144                     With rsCashOnBank
146                         .AddNew
148                         !cashonbank = Format$(Val(txtCashonbank.Text), _
                                    "###,###,##0.00")
150                         !UpdateDate = lblDate.Caption
152                         !User = lblUser.Caption
154                         !TransactionDate = dtOnBank.Value
156                         .Update
                        End With

158                     DataGrid2.Refresh
160                     Unload Me
162                     frm_CashOnHand.lblUser.Caption = MDIForm1.lblUserName.Caption
164                     frm_CashOnHand.Show
                    End If
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnEditBank_Click"

        Exit Sub

btnEditBank_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.btnEditBank_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEditCashB_Click()

        '<EhHeader>
        On Error GoTo btnEditCashB_Click_Err

        TxtLog "Entered btnEditCashB_Click"

        '</EhHeader>

100     If btnEditCashB.Caption = "&Edit Cash On Bank" Then
102         btnEditCashB.Caption = "&Update"
104         btnClose.Caption = "&Cancel"
106         btnEdit.Enabled = False
108         btnClose.Caption = "&Cancel"
110         txtCashonbank.Enabled = True
112         txtCashonbank.SetFocus
114         btnClose1.Caption = "&Cancel"
116     ElseIf btnEditCashB.Caption = "&Update" Then

118         If txtCashonbank.Text = "" Then
120             MsgBox "Field must not be blank.", vbInformation, _
                        "Webplus Lending Corporation"
122             txtCashonbank.SetFocus
124             Call cashonbank
126             Set DataGrid2.DataSource = rsCashOnBank
            Else
128             Print rsCashOnBank!ID
                    
                Dim tempCOHID As Integer

                'tempCOHID = rsCashOnBank!ID
130             tempCOHID = lblIDb.Caption

132             If rsCashOnBank.State = 1 Then rsCashOnBank.Close
                    
134             rsCashOnBank.Open _
                        "Select * from tblCashOnBank where Transactiondate = #" & _
                        dtOnBank.Value & "# and ID <> " & tempCOHID
                    
136             If rsCashOnBank.RecordCount <> 0 Then
138                 MsgBox "Cash on hand is already encoded on this date.", vbInformation
140                 txtCashonbank.SetFocus

142                 If rsCashOnBank.State = 1 Then rsCashOnBank.Close
144                 rsCashOnBank.Open "Select * from tblCashOnBank Order by ID desc"
146                 Set DataGrid2.DataSource = rsCashOnBank
                Else

148                 If MsgBox("Are you sure you want to update cash on Bank balance?", _
                            vbQuestion + vbYesNo) = vbYes Then

150                     If rsCashOnBank.State = 1 Then rsCashOnBank.Close
152                     rsCashOnBank.Open "Select * from tblCashOnBank where ID = " & _
                                Val(lblIDb.Caption) & ""

154                     If rsCashOnBank.RecordCount <> 0 Then

156                         With rsCashOnBank
158                             !cashonbank = txtCashonbank.Text
160                             !UpdateDate = lblDate.Caption
162                             !User = lblUser.Caption
164                             !TransactionDate = dtOnBank.Value 'zzzz
166                             .Update
                            End With

168                         DataGrid2.Refresh
170                         Unload Me
   
172                         frm_CashOnHand.lblUser.Caption = MDIForm1.lblUserName.Caption
174                         frm_CashOnHand.Show
                        End If

                    Else

176                     If rsCashOnBank.State = 1 Then rsCashOnBank.Close
178                     rsCashOnBank.Open "Select * from tblCashOnBank Order by ID desc"
180                     Set DataGrid2.DataSource = rsCashOnBank
                    End If
                End If
            End If
   
        End If

        '<EhFooter>

        TxtLog "Exited btnEditCashB_Click"

        Exit Sub

btnEditCashB_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.btnEditCashB_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdEditCash_Click()

        '<EhHeader>
        On Error GoTo cmdEditCash_Click_Err

        TxtLog "Entered cmdEditCash_Click"

        '</EhHeader>

100     If cmdEditCash.Caption = "&Edit Cash on Hand" Then
102         cmdEditCash.Caption = "&Update"
104         btnClose.Caption = "&Cancel"
106         txtCashOnHand.Enabled = True
108         txtCashOnHand.SetFocus
110         btnEdit.Enabled = False
112         btnClose.Caption = "&Cancel"
114         btnClose1.Caption = "&Cancel"
116     ElseIf cmdEditCash.Caption = "&Update" Then

118         If txtCashOnHand.Text = "" Then
120             MsgBox "Field must not be blank.", vbInformation, _
                        "Webplus Lending Corporation"
122             txtCashOnHand.SetFocus
124             Call Cashonhand
126             Set DataGrid1.DataSource = rsCashOnHand
            Else
128             Print rsCashOnHand!ID
                    
                Dim tempCOHID As Integer

130             tempCOHID = rsCashOnHand!ID

132             If rsCashOnHand.State = 1 Then rsCashOnHand.Close
                    
134             rsCashOnHand.Open _
                        "Select * from tblCashOnHand where Transactiondate = #" & _
                        dtOnHand.Value & "# and ID <> " & tempCOHID
                    
136             If rsCashOnHand.RecordCount <> 0 Then
138                 MsgBox _
                            "Cash on hand is already encoded on this date.Just edit if you wish to change the value.", _
                            vbInformation
140                 txtCashOnHand.SetFocus

142                 If rsCashOnHand.State = 1 Then rsCashOnHand.Close
144                 rsCashOnHand.Open "Select * from tblCashOnHand Order by ID desc"
146                 Set DataGrid1.DataSource = rsCashOnHand
        
                Else

148                 If MsgBox("Are you sure you want to update cash on hand balance?", _
                            vbQuestion + vbYesNo) = vbYes Then
                 
150                     If rsCashOnHand.State = 1 Then rsCashOnHand.Close
152                     rsCashOnHand.Open "Select * from tblCashOnHand where ID = " & _
                                Val(lblID.Caption) & ""

154                     If rsCashOnHand.RecordCount <> 0 Then
                    
156                         With rsCashOnHand
                            
158                             !Cashonhand = Format$(Val(txtCashOnHand.Text), _
                                        "###,###,##0.00")
160                             !UpdateDate = lblDate.Caption
162                             !User = lblUser.Caption
164                             !TransactionDate = dtOnHand.Value
166                             .Update
                            End With

168                         If rsCashOnHand.State = 1 Then rsCashOnHand.Close
170                         rsCashOnHand.Open _
                                    "Select * from tblCashOnHand order by ID desc"
172                         DataGrid1.Refresh
174                         Unload Me
                      
176                         frm_CashOnHand.lblUser.Caption = MDIForm1.lblUserName.Caption
178                         frm_CashOnHand.Show
                        End If

                    Else

180                     If rsCashOnHand.State = 1 Then rsCashOnHand.Close
182                     rsCashOnHand.Open "Select * from tblCashOnHand Order by ID desc"
184                     Set DataGrid1.DataSource = rsCashOnHand
                    End If
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited cmdEditCash_Click"

        Exit Sub

cmdEditCash_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.cmdEditCash_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        TxtLog "Entered DataGrid1_Click"

        '</EhHeader>

100     If rsCashOnHand.RecordCount = 0 Then
        Else
102         txtCashOnHand.Text = rsCashOnHand!Cashonhand
104         lblOnHand.Caption = rsCashOnHand!TransactionDate
106         dtOnHand.Value = rsCashOnHand!TransactionDate
108         lblID.Caption = rsCashOnHand!ID
            'btnClose.Caption = "&Cancel"
110         btnClose1.Caption = "&Cancel"
112         DataGrid2.Enabled = False
114         btnEdit.Enabled = False
116         dtOnHand.Enabled = True
118         cmdEditCash.Enabled = True
        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_Click"

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.DataGrid1_Click", Erl

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
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid2_Click()

        '<EhHeader>
        On Error GoTo DataGrid2_Click_Err

        TxtLog "Entered DataGrid2_Click"

        '</EhHeader>

100     If rsCashOnBank.RecordCount = 0 Then
        Else
102         txtCashonbank.Text = rsCashOnBank!cashonbank
104         lblOnBank.Caption = rsCashOnBank!TransactionDate
106         dtOnBank.Value = rsCashOnBank!TransactionDate
108         lblIDb.Caption = rsCashOnBank!ID
110         btnEdit.Enabled = False
112         btnClose.Caption = "&Cancel"
            ' btnClose1.Caption = "&Cancel"
114         DataGrid1.Enabled = False
116         btnEditBank.Enabled = False
118         btnEditCashB.Enabled = True
        End If

        '<EhFooter>

        TxtLog "Exited DataGrid2_Click"

        Exit Sub

DataGrid2_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.DataGrid2_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid2_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DataGrid2_KeyPress_Err

        TxtLog "Entered DataGrid2_KeyPress"

        '</EhHeader>

100     KeyAscii = 0

        '<EhFooter>

        TxtLog "Exited DataGrid2_KeyPress"

        Exit Sub

DataGrid2_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.DataGrid2_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call Cashonhand
104     Call cashonbank
        
106     If rsCashOnHand.State = 1 Then rsCashOnHand.Close
108     rsCashOnHand.Open "Select * from tblCashOnHand Order by ID desc"
110     Set DataGrid1.DataSource = rsCashOnHand
112     DataGrid1.Refresh

114     If rsCashOnBank.State = 1 Then rsCashOnBank.Close
116     rsCashOnBank.Open "Select * from tblCashOnBank Order by ID desc"
118     Set DataGrid2.DataSource = rsCashOnBank

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_CashOnHand.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    lblDate.Caption = Date
    dtOnHand.Value = Date
    dtOnBank.Value = Date
    
    Timer1.Enabled = False
End Sub

Private Sub txtCashonbank_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtCashonbank_KeyPress_Err

        TxtLog "Entered txtCashonbank_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321."

        'Check for Only One Decimal Point In Textbox
102     If KeyAscii = 46 Then

            'If more than one decimal point is typed, only one decimal will be printed
104         If InStr(1, txtCashonbank.Text, ".") > 0 Then
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
118         KeyAscii = 13
        Else
120         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtCashonbank_KeyPress"

        Exit Sub

txtCashonbank_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CashOnHand.txtCashonbank_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtCashOnHand_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtCashOnHand_KeyPress_Err

        TxtLog "Entered txtCashOnHand_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321."

        'Check for Only One Decimal Point In Textbox
102     If KeyAscii = 46 Then

            'If more than one decimal point is typed, only one decimal will be printed
104         If InStr(1, txtCashOnHand.Text, ".") > 0 Then
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
118         KeyAscii = 13
     
        Else
120         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtCashOnHand_KeyPress"

        Exit Sub

txtCashOnHand_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CashOnHand.txtCashOnHand_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

