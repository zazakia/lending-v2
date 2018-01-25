VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Chart 
   Caption         =   "frmChart"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10320
      Top             =   480
   End
   Begin VB.Frame frEdit 
      BackColor       =   &H00FF8080&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   10680
      TabIndex        =   28
      Top             =   1800
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtAcID 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2040
         TabIndex        =   45
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   735
         Left            =   3000
         TabIndex        =   33
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Edit"
         Height          =   735
         Left            =   480
         TabIndex        =   32
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtAmt 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   31
         Text            =   "0"
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox txtDesc 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   30
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtAccntName 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   29
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label tetrstrt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account ID   :"
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
         TabIndex        =   44
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount              :"
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
         TabIndex        =   37
         Top             =   2280
         Width           =   1740
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description        :"
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
         TabIndex        =   36
         Top             =   1800
         Width           =   1740
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name  :"
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
         TabIndex        =   35
         Top             =   1320
         Width           =   1755
      End
   End
   Begin VB.Frame frDescription 
      BackColor       =   &H00FF8080&
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   5400
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cbActype 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtAcctName2 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   405
         Left            =   1920
         TabIndex        =   39
         Top             =   1440
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   25
         Top             =   2880
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   265748481
         CurrentDate     =   41871
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Close"
         Height          =   615
         Left            =   2640
         TabIndex        =   22
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Height          =   615
         Left            =   240
         TabIndex        =   21
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Text            =   "0"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtDescription 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name:  "
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
         TabIndex        =   38
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "To Date          :"
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
         TabIndex        =   24
         Top             =   2880
         Width           =   1515
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Amount          :"
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
         TabIndex        =   20
         Top             =   2400
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Type                :"
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
         TabIndex        =   18
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Account No    :"
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
         Top             =   480
         Width           =   1545
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Description    :"
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
         TabIndex        =   16
         Top             =   1920
         Width           =   1500
      End
   End
   Begin VB.Frame frAdd 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtDescription1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   47
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton btnClose1 
         Caption         =   "&Close"
         Height          =   615
         Left            =   2400
         TabIndex        =   23
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton btn1 
         Caption         =   "&Add"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txtAccountName 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cbType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frm_Chart.frx":0000
         Left            =   2040
         List            =   "frm_Chart.frx":0002
         TabIndex        =   4
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description        :"
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
         TabIndex        =   48
         Top             =   2520
         Width           =   1740
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name  :"
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
         TabIndex        =   10
         Top             =   2040
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type                    :"
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
         Top             =   1560
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account No        :"
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
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add  an Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1080
         TabIndex        =   7
         Top             =   360
         Width           =   2580
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chart of Accounts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   735
         Left            =   6960
         TabIndex        =   40
         Top             =   8640
         Width           =   1935
      End
      Begin VB.CommandButton btnEditt 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   735
         Left            =   3360
         TabIndex        =   34
         Top             =   8640
         Width           =   2055
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   2040
         TabIndex        =   26
         Top             =   600
         Width           =   8055
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   8640
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7455
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   13150
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Search Here:"
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
         TabIndex        =   27
         Top             =   600
         Width           =   1410
      End
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   9960
      TabIndex        =   43
      Top             =   3720
      Width           =   45
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   10080
      TabIndex        =   42
      Top             =   3120
      Width           =   45
   End
   Begin VB.Label lblType 
      Height          =   255
      Left            =   15840
      TabIndex        =   41
      Top             =   1920
      Width           =   495
   End
End
Attribute VB_Name = "frm_Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Chart
'    Project    : Project1
'
'    Description: [This module will Create a Chart Accounts , could be Income or Expense]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Sub autoNumber()

        '<EhHeader>
        On Error GoTo autoNumber_Err

        TxtLog "Entered autoNumber"

        '</EhHeader>

100     If rsChart.RecordCount = 0 Then
102         txtAccountNo.Text = "AC00001"
        Else
104         rsChart.MoveLast
106         txtAccountNo.Text = "AC" & Format$(Right$(rsChart!AccountNo, 5) + 1, "00000")

        End If

        '<EhFooter>

        TxtLog "Exited autoNumber"

        Exit Sub

autoNumber_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.autoNumber", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btn1_Click()

        '<EhHeader>
        On Error GoTo btn1_Click_Err

        TxtLog "Entered btn1_Click"

        '</EhHeader>

100     If btn1.Caption = "&Add" Then
102         btn1.Caption = "&Save"
104         btnClose1.Caption = "&Cancel"
106         cbType.Enabled = True
        
108         cbType.SetFocus
110         txtAccountName.Enabled = True
112         txtDescription1.Enabled = True
    
114     ElseIf txtAccountName.Text = "" Or txtDescription1.Text = "" Then
116         MsgBox "All fields are required", vbInformation, "J Lending"
118         txtAccountName.SetFocus

        Else

120         If rsChart.State = 1 Then rsChart.Close
122         rsChart.Open "Select * from tblChartOfAccounts where AccountName = '" & _
                    Trim$(txtAccountName.Text) & "' and Description = '" & Trim$( _
                    txtDescription1.Text) & "' and Type = '" & cbType.Text & "'"

124         If rsChart.RecordCount <> 0 Then
126             MsgBox "Record Already exist!", vbInformation, "J Lending Corp"
128             Call Chart
130             Set DataGrid1.DataSource = rsChart
            Else

132             If MsgBox("Are you sure you want to add an Account?", vbQuestion + _
                        vbYesNo) = vbYes Then

134                 If rsChart.State = 1 Then rsChart.Close
136                 rsChart.Open _
                            "Select * from tblChartOfAccounts where AccountName = '" & _
                            Trim$(txtAccountName.Text) & "'"

138                 With rsChart
140                     .AddNew
142                     !AccountNo = txtAccountNo.Text
144                     !Type = cbType.Text
146                     !AccountName = Trim$(txtAccountName.Text)
148                     !Description = Trim$(txtDescription1.Text)
150                     .Update
                    End With

152                 MsgBox "Record Successfully Added", vbInformation, "J Lending Corp"
154                 Unload Me
156                 Me.Show
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btn1_Click"

        Exit Sub

btn1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.btn1_Click", Erl

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
104         frAdd.Visible = True
        Else
    
        End If

        '<EhFooter>

        TxtLog "Exited btnAdd_Click"

        Exit Sub

btnAdd_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.btnAdd_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnClose1_Click()

        '<EhHeader>
        On Error GoTo btnClose1_Click_Err

        TxtLog "Entered btnClose1_Click"

        '</EhHeader>

100     If btnClose1.Caption = "&Close" Then
102         txtAccountName.Text = ""
104         frAdd.Visible = False
106         btnAdd.Caption = "&Add"
        
        Else
108         txtAccountName.Text = ""
110         btnClose1.Caption = "&Close"
112         btn1.Caption = "&Add"
        
        End If

        '<EhFooter>

        TxtLog "Exited btnClose1_Click"

        Exit Sub

btnClose1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.btnClose1_Click", Erl

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
108         btnAdd.Enabled = True
110         btnEdit.Enabled = False
112         btnEditt.Enabled = False
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.btnClose_Click", Erl

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
104         Command3.Caption = "&Cancel"
106         txtDescription.Enabled = True
108         txtAcctName2.Enabled = True
110         cbActype.Enabled = True
112         txtDescription.SetFocus
            
        Else
            
114         If cbActype.Text = "" Or txtAcctName2.Text = "" Or txtDescription.Text = "" _
                    Or txtAmount.Text = "" Then
116             MsgBox "Add fields are required.", vbInformation, _
                        "Lending Webplus Corporation"
        
            Else

118             If rsChart.State = 1 Then rsChart.Close
120             rsChart.Open "Select * from tblChartOfAccounts where AccountName = '" & _
                        Trim$(txtAcctName2.Text) & "' and Description = '" & Trim$( _
                        txtDescription.Text) & "' and Type = '" & cbActype.Text & _
                        "' and AccountNo <> '" & Text2.Text & "'"
            
122             If rsChart.RecordCount <> 0 Then
124                 MsgBox "Cannot edit records to equal another record.", _
                            vbInformation, "J Lending Corp"
126                 Call Chart
128                 Set DataGrid1.DataSource = rsChart
            
                Else
                    ''start
        
130                 If rsChart.State = 1 Then rsChart.Close
132                 rsChart.Open _
                            " Select * from tblChartOfAccounts where AccountNo = '" & _
                            Text2.Text & "'"
            
134                 With rsChart
136                     !AccountNo = Text2.Text
138                     !Type = cbActype.Text
140                     !AccountName = Trim$(txtAcctName2.Text)
142                     !Description = Trim$(txtDescription.Text)
144                     !Amount = txtAmount.Text
146                     !ToDate = DTPicker1.Value
148                     .Update
                    End With
            
                    '            Dim e As Integer
                    '                If rsExpense.State = 1 Then rsExpense.Close
                    '                rsExpense.Open "Select * from tblExpense where AccountName = '" & Text2.Text & "'"
                    '                If rsExpense.RecordCount <> 0 Then
                    '                rsExpense.MoveFirst
                    '                For e = 0 To rsExpense.RecordCount
                    '                    rsExpense.Close
                    '                     rsExpense.Open "Select * from tblExpense where AccountName = '" & Text2.Text & "'"
                    '                If rsExpense.RecordCount <> 0 Then
                    '                    rsExpense!AccountName = Trim$(txtAcctName2.Text)
                    '                    rsExpense.Update
                    '                    End If
                    '                    Next e
                    '                End If
                    '
                    '                Dim d As Integer
                    '                If rsDeposit.State = 1 Then rsDeposit.Close
                    '                rsDeposit.Open "Select * from tblDeposit where AccountName = '" & Text2.Text & "'"
                    '                If rsDeposit.RecordCount <> 0 Then
                    '                rsDeposit.MoveFirst
                    '                For d = 0 To rsDeposit.RecordCount
                    '                    rsDeposit.Close
                    '                    rsDeposit.Open "Select * from tblDeposit where AccountName = '" & Text2.Text & "'"
                    '                If rsDeposit.RecordCount <> 0 Then
                    '                    rsDeposit!AccountName = Trim$(txtAcctName2.Text)
                    '                    rsDeposit.Update
                    '                    End If
                    '                    Next d
                    '                End If

150                 MsgBox "Record Successfully Updated", vbInformation, "J Lending Corp"
152                 Unload Me
154                 Me.Show
        
                    ''end
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnEdit_Click"

        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.btnEdit_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEditt_Click()

        '<EhHeader>
        On Error GoTo btnEditt_Click_Err

        TxtLog "Entered btnEditt_Click"

        '</EhHeader>

        '    DataGrid1.Refresh
100     If rsChart.RecordCount = 0 Then
            
        Else
            ' btnEditt.Enabled = True
102         Text2.Text = rsChart!AccountNo
104         txtAcctName2.Text = rsChart!AccountName
106         cbActype.Text = rsChart!Type

108         If frDescription.Visible = False Then
110             frDescription.Visible = True
            End If

            '102         frEdit.Visible = True
            '            If frDescription.Visible = True Then
            '                frDescription.Visible = False
            '            With rsChart
            '                txtAcID.Text = !AccountNo
            '104             txtAccntName.Text = !AccountName
            '106             txtDesc.Text = ""
            '            End With
            '            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnEditt_Click"

        Exit Sub

btnEditt_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.btnEditt_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cbType_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo cbType_KeyPress_Err

        TxtLog "Entered cbType_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         txtAccountName.SetFocus
        Else
104         KeyAscii = 0
    
        End If

        '<EhFooter>

        TxtLog "Exited cbType_KeyPress"

        Exit Sub

cbType_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.cbType_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        TxtLog "Entered Command1_Click"

        '</EhHeader>

100     If Command1.Caption = "&Edit" Then
102         Command1.Caption = "&Update"
104         Command2.Caption = "&Cancel"
106         txtAccntName.Enabled = True
108         txtDesc.Enabled = True
110     ElseIf Command1.Caption = "&Update" Then

112         If rsChart.State = 1 Then rsChart.Close
114         rsChart.Open " Select * from tblChartOfAccounts where AccountNo = '" & _
                    txtAcID.Text & "'"

116         If rsChart.RecordCount = 0 Then
118             MsgBox "No Record Found", vbInformation
            Else

120             With rsChart
122                 !AccountNo = txtAcID.Text
124                 !AccountName = Trim$(txtAccntName.Text)
126                 !Description = Trim$(txtDesc.Text)
128                 !Amount = txtAmt.Text
130                 .Update
                End With

                Dim e As Integer

132             If rsExpense.State = 1 Then rsExpense.Close
134             rsExpense.Open "Select * from tblExpense where AccountName = '" & _
                        lblName.Caption & "'"

136             If rsExpense.RecordCount <> 0 Then
138                 rsExpense.MoveFirst

140                 For e = 0 To rsExpense.RecordCount
142                     rsExpense.Close
144                     rsExpense.Open "Select * from tblExpense where AccountName = '" _
                                & lblName.Caption & "'"

146                     If rsExpense.RecordCount <> 0 Then
148                         rsExpense!AccountName = Trim$(txtAccntName.Text)
150                         rsExpense.Update
                        End If

152                 Next e

                End If
                
                Dim d As Integer

154             If rsDeposit.State = 1 Then rsDeposit.Close
156             rsDeposit.Open "Select * from tblDeposit where AccountName = '" & _
                        lblName.Caption & "'"

158             If rsDeposit.RecordCount <> 0 Then
160                 rsDeposit.MoveFirst

162                 For d = 0 To rsDeposit.RecordCount
164                     rsDeposit.Close
166                     rsDeposit.Open "Select * from tblDeposit where AccountName = '" _
                                & lblName.Caption & "'"

168                     If rsDeposit.RecordCount <> 0 Then
170                         rsDeposit!AccountName = txtAccntName.Text
172                         rsDeposit.Update
                        End If

174                 Next d

                End If

176             MsgBox "Record Successfuly Edited.", vbInformation
178             Unload Me
180             Me.Show
                
            End If
    
        End If

        '<EhFooter>

        TxtLog "Exited Command1_Click"

        Exit Sub

Command1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.Command1_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command2_Click()

        '<EhHeader>
        On Error GoTo Command2_Click_Err

        TxtLog "Entered Command2_Click"

        '</EhHeader>

100     If Command2.Caption = "&Close" Then
102         frEdit.Visible = False
        Else
104         Command1.Caption = "&Edit"
106         Command2.Caption = "&Close"
108         txtAccntName.Enabled = False
110         txtDesc.Enabled = False
112         txtAmt.Enabled = False
        End If

        '<EhFooter>

        TxtLog "Exited Command2_Click"

        Exit Sub

Command2_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.Command2_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Command3_Click()

        '<EhHeader>
        On Error GoTo Command3_Click_Err

        TxtLog "Entered Command3_Click"

        '</EhHeader>

100     cbActype.Enabled = False
102     txtAcctName2.Enabled = False
104     txtDescription.Enabled = False
        
106     frDescription.Visible = False

        '<EhFooter>

        TxtLog "Exited Command3_Click"

        Exit Sub

Command3_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.Command3_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        TxtLog "Entered DataGrid1_Click"

        '</EhHeader>

100     btnEditt.Enabled = True
102     btnAdd.Enabled = False
104     btnClose.Caption = "&Cancel"
106     lblID.Caption = rsChart!AccountNo
108     lblName.Caption = rsChart!AccountName

110     If frEdit.Visible = True Then
112         frEdit.Visible = False
        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_Click"

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.DataGrid1_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_DblClick()

        '<EhHeader>
        On Error GoTo DataGrid1_DblClick_Err

        TxtLog "Entered DataGrid1_DblClick"

        '</EhHeader>

100     If rsChart.RecordCount = 0 Then
        Else
102         frDescription.Visible = True
    
104         btnAdd.Enabled = False

106         With rsChart
108             Text2.Text = !AccountNo
110             cbActype.Text = !Type
112             txtAcctName2.Text = !AccountName
        
            End With

        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_DblClick"

        Exit Sub

DataGrid1_DblClick_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.DataGrid1_DblClick", Erl

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
        ErrReport Err.Description, "LendingClient.frm_Chart.DataGrid1_KeyPress", Erl

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
104     Call autoNumber
106     Call Expense
108     Call Deposit
110     DataGrid1.Refresh
112     Set DataGrid1.DataSource = rsChart
114     DataGrid1.Refresh
        
116     cbType.AddItem "Income"
118     cbType.AddItem "Expense"
        
120     cbActype.AddItem "Income"
122     cbActype.AddItem "Expense"

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    DTPicker1.Value = Date
    Timer1.Enabled = False

End Sub

Private Sub txtAccountName_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtAccountName_KeyPress_Err

        TxtLog "Entered txtAccountName_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz "

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then

104     ElseIf KeyAscii = 13 Then

106         txtDescription1.SetFocus
    
108     ElseIf KeyAscii = 8 Then
110         KeyAscii = 8

        Else
112         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtAccountName_KeyPress"

        Exit Sub

txtAccountName_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.txtAccountName_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtAmount_KeyPress_Err

        TxtLog "Entered txtAmount_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321."

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf KeyAscii = 13 Then
110         DTPicker1.SetFocus
         
        Else
112         KeyAscii = 0
        
        End If

        '<EhFooter>

        TxtLog "Exited txtAmount_KeyPress"

        Exit Sub

txtAmount_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.txtAmount_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtDescription1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtDescription1_KeyPress_Err

        TxtLog "Entered txtDescription1_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         btn1.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtDescription1_KeyPress"

        Exit Sub

txtDescription1_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.txtDescription1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtDescription_KeyPress_Err

        TxtLog "Entered txtDescription_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         txtAmount.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtDescription_KeyPress"

        Exit Sub

txtDescription_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.txtDescription_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsChart.State = 1 Then rsChart.Close
102     rsChart.Open "Select * from tblChartOfAccounts where AccountName like '" & _
                txtSearch.Text & "%'"
104     Set DataGrid1.DataSource = rsChart

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Chart.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

