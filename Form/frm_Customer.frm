VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Customer 
   Caption         =   "frmCustomer"
   ClientHeight    =   11055
   ClientLeft      =   28740
   ClientTop       =   4065
   ClientWidth     =   20130
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   20130
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   10080
      TabIndex        =   31
      Top             =   10200
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   261423105
      CurrentDate     =   41878
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7320
      Top             =   1800
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20055
      Begin VB.TextBox txtCFirstname 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   9360
         TabIndex        =   46
         Top             =   3480
         Width           =   2895
      End
      Begin VB.TextBox txtCLastname 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   9360
         TabIndex        =   45
         Top             =   3000
         Width           =   2895
      End
      Begin VB.CommandButton btnReverse 
         Caption         =   "&Reverse"
         Enabled         =   0   'False
         Height          =   465
         Left            =   4200
         TabIndex        =   40
         Top             =   8520
         Width           =   1455
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   17280
         TabIndex        =   35
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   2400
         TabIndex        =   27
         Top             =   4560
         Width           =   5055
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3375
         Left            =   120
         TabIndex        =   26
         Top             =   5040
         Width           =   19815
         _ExtentX        =   34951
         _ExtentY        =   5953
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
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2280
         TabIndex        =   25
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   495
         Left            =   6120
         TabIndex        =   18
         Top             =   8520
         Width           =   1455
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   8520
         Width           =   1455
      End
      Begin VB.TextBox txtRemarks 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   405
         Left            =   2040
         TabIndex        =   16
         Top             =   3960
         Width           =   4215
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   9360
         TabIndex        =   14
         Text            =   "0"
         Top             =   3960
         Width           =   2895
      End
      Begin VB.ComboBox cbCollector 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   3120
         Width           =   4215
      End
      Begin VB.TextBox txtMI 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   8
         Top             =   2640
         Width           =   4215
      End
      Begin VB.TextBox txtFirstname 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtLastname 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   7080
         TabIndex        =   54
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblRequiredField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Required Field"
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
         Left            =   7560
         TabIndex        =   53
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label10 
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
         Left            =   5400
         TabIndex        =   52
         Top             =   2160
         Width           =   90
      End
      Begin VB.Label lblLabel110 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instructions: All required fields must not  be blank"
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
         Left            =   360
         TabIndex        =   51
         Top             =   360
         Width           =   4830
      End
      Begin VB.Label lblLabel19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH USING LAST NAME , FIRST NAME ,  OR CODE"
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
         Left            =   7560
         TabIndex        =   50
         Top             =   4560
         Width           =   6255
      End
      Begin VB.Label lblAuCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "vvvv"
         Height          =   315
         Left            =   12840
         TabIndex        =   49
         Top             =   3960
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblCollectorFirst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Collector First Name : "
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
         Left            =   6840
         TabIndex        =   48
         Top             =   3480
         Width           =   2325
      End
      Begin VB.Label lblccccccc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Collector LastName :"
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
         Left            =   6840
         TabIndex        =   47
         Top             =   3000
         Width           =   2205
      End
      Begin VB.Label lblLName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   13920
         TabIndex        =   44
         Top             =   840
         Width           =   45
      End
      Begin VB.Label lblCollectorFname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   14040
         TabIndex        =   43
         Top             =   360
         Width           =   45
      End
      Begin VB.Label lblCollectorCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "coo"
         Height          =   195
         Left            =   8280
         TabIndex        =   42
         Top             =   720
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblCustomerS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note : Customer's full name should only be typed by letters only. "
         Height          =   195
         Left            =   1560
         TabIndex        =   41
         Top             =   1320
         Width           =   4515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "User : "
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
         Left            =   16440
         TabIndex        =   36
         Top             =   240
         Width           =   690
      End
      Begin VB.Line Line1 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   120
         X2              =   20040
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Search here       :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   28
         Top             =   4560
         Width           =   2160
      End
      Begin VB.Label jdfjdf 
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
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   23
         Top             =   3600
         Width           =   255
      End
      Begin VB.Label Label14 
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
         Left            =   6240
         TabIndex        =   22
         Top             =   2640
         Width           =   90
      End
      Begin VB.Label Label13 
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
         Left            =   6240
         TabIndex        =   21
         Top             =   3120
         Width           =   135
      End
      Begin VB.Label Label12 
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
         Left            =   6240
         TabIndex        =   20
         Top             =   3960
         Width           =   90
      End
      Begin VB.Label Label11 
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
         Left            =   5400
         TabIndex        =   19
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Remarks          :"
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
         TabIndex        =   15
         Top             =   3960
         Width           =   1620
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Balance                     :"
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
         Left            =   6840
         TabIndex        =   13
         Top             =   3960
         Width           =   2190
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Collector           :"
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
         TabIndex        =   12
         Top             =   3600
         Width           =   1650
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Address            :"
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
         Top             =   3120
         Width           =   1665
      End
      Begin VB.Label Label4 
         Caption         =   "Middle Initial     :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "First Name       :"
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
         TabIndex        =   3
         Top             =   2160
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Last Name       :"
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
         TabIndex        =   2
         Top             =   1680
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code                 :"
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
         TabIndex        =   1
         Top             =   840
         Width           =   1650
      End
   End
   Begin VB.Label lblCollector 
      Caption         =   "Label19"
      Height          =   255
      Left            =   10680
      TabIndex        =   39
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      Height          =   255
      Left            =   10800
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblDate 
      Caption         =   "Label15"
      Height          =   15
      Left            =   10680
      TabIndex        =   37
      Top             =   480
      Width           =   135
   End
   Begin VB.Label lbluser 
      Caption         =   "lbluser"
      Height          =   375
      Left            =   10080
      TabIndex        =   24
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label lblFirstname 
      AutoSize        =   -1  'True
      Caption         =   "FFF"
      Height          =   195
      Left            =   10680
      TabIndex        =   34
      Top             =   1920
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblLastname 
      AutoSize        =   -1  'True
      Caption         =   "Label7"
      Height          =   195
      Left            =   10680
      TabIndex        =   33
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label lblUserlevel 
      Caption         =   "Label7"
      Height          =   255
      Left            =   10680
      TabIndex        =   32
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label lblTime 
      Height          =   255
      Left            =   10680
      TabIndex        =   30
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label lblCode 
      Caption         =   "lblCode"
      Height          =   255
      Left            =   10680
      TabIndex        =   29
      Top             =   3720
      Width           =   135
   End
End
Attribute VB_Name = "frm_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Customer
'    Project    : Project1
'
'    Description: [This procedure will add a New Customer Record]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Sub addItems()

        '<EhHeader>
        On Error GoTo addItems_Err

        TxtLog "Entered addItems"

        '</EhHeader>

        '
        
        '        If rsCollCode.State = 1 Then rsCollCode.Close
        '        rsCollCode.Open "Select * from tblColl_Code"
        '        cbCollector.Clear
        '
        '        If rsCollCode.RecordCount <> 0 Then
        '
        '            Do While Not rsCollCode.EOF
        '                cbCollector.AddItem rsCollCode!code
        '                rsCollCode.MoveNext
        '            Loop
        '
        '        End If
100     If rsCollData.State = 1 Then rsCollData.Close
102     rsCollData.Open "Select * from tblColl_Data"
104     cbCollector.Clear

106     If rsCollData.RecordCount <> 0 Then

108         Do While Not rsCollData.EOF
110             cbCollector.AddItem rsCollData!code
112             rsCollData.MoveNext
            Loop

        End If

        '<EhFooter>

        TxtLog "Exited addItems"

        Exit Sub

addItems_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.addItems", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub autoNumber()

        '<EhHeader>
        On Error GoTo autoNumber_Err

        TxtLog "Entered autoNumber"

        '</EhHeader>

100     If rsCustomer.RecordCount = 0 Then
102         txtCode.Text = "00001"
104         lblAuCode.Caption = "00001"
        Else
106         DataGrid1.Refresh
108         rsCustomer.MoveFirst
110         txtCode.Text = "" & Format$(Right$(rsCustomer!code, 5) + 1, "00000")
112         DataGrid1.Refresh
114         lblAuCode.Caption = txtCode.Text
        End If

        '<EhFooter>

        TxtLog "Exited autoNumber"

        Exit Sub

autoNumber_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.autoNumber", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnAdd_Click()

        '<EhHeader>
        On Error GoTo btnAdd_Click_Err

        TxtLog "Entered btnAdd_Click"

        '</EhHeader>

100     txtSearch.Enabled = False

102     If btnAdd.Caption = "&Add" Then
104         If rsCustomer.State = 1 Then rsCustomer.Close
            'Sort records descending
106         rsCustomer.Open "Select * from tblCustomer Order By Code desc"
108         Call autoNumber

110         Set DataGrid1.DataSource = rsCustomer
            'Adjusting the width of Fields on Datagird
112         DataGrid1.Width = Me.Width
114         DataGrid1.Columns(0).Width = 0
116         DataGrid1.Columns(1).Width = 550
118         DataGrid1.Columns(4).Width = 950
120         DataGrid1.Columns(5).Width = 4000
122         DataGrid1.Columns(7).Width = 2000
124         DataGrid1.Columns(8).Width = 1100

126         btnAdd.Caption = "&Save"
128         btnClose.Caption = "&Cancel"
    
130         txtLastname.Enabled = True
132         txtFirstname.Enabled = True
134         txtMI.Enabled = True
136         txtAddress.Enabled = True
138         cbCollector.Enabled = True
140         txtRemarks.Enabled = True
142         DataGrid1.Enabled = False
   
144         txtLastname.SetFocus
    
        Else
            
            'Check if any required fields is blank.
146         If Trim$(txtLastname.Text) = "" Or Trim$(txtFirstname.Text) = "" Or Trim$( _
                    txtMI.Text) = "" Or Trim$(txtAddress.Text) = "" Or cbCollector.Text _
                    = "" Or Trim$(txtRemarks.Text) = "" Then
                'If there is a blank in any required fields . This message box will display.
148             MsgBox "All fields are required.", vbInformation, _
                        "Webplus Lending Corporation"
            Else
                'Ask if the you are ready to save this record to database.
        
150             If rsCustomer.State = 1 Then rsCustomer.Close
152             rsCustomer.Open "Select * from tblCustomer where Lastname = '" & _
                        txtLastname.Text & "' and Firstname = '" & txtFirstname.Text & _
                        "'"

154             If rsCustomer.RecordCount <> 0 Then
                    'If there is a duplicate record. This message box will show
156                 MsgBox "Account already exist", vbInformation, _
                            "Webplus Lending Corporation"
158                 txtFirstname.SetFocus
                Else

                    'New Record is saved to database
160                 If rsCustomer.State = 1 Then rsCustomer.Close
                    'Sort records descending
162                 rsCustomer.Open "Select * from tblCustomer Order By Code desc"

164                 If MsgBox("Are you sure you want to add new Record?", vbQuestion + _
                            vbYesNo, "Webplus Lending Corporation") = vbYes Then
166                     DataGrid1.Refresh
                        
168                     Call autoNumber

170                     With rsCustomer
172                         .AddNew
174                         !code = txtCode.Text
176                         !lastname = Trim$(txtLastname.Text)
178                         !firstname = Trim$(txtFirstname.Text)
180                         !MiddleInitial = txtMI.Text
182                         !Address = Trim$(txtAddress.Text)
184                         !CollectorCode = cbCollector.Text
186                         !Collector = txtCLastname.Text
188                         !Amortization = "0.00"
190                         !Balance = "0"
192                         !Remarks = Trim$(txtRemarks.Text)
194                         !User = txtUser.Text
196                         !CollectorFirstname = txtCFirstname.Text
198                         .Update
                        End With

200                     DataGrid1.Refresh

202                     If rsCustomer.State = 1 Then rsCustomer.Close
                        'Sort records descending
204                     rsCustomer.Open "Select * from tblCustomer Order By Code desc"
206                     Call autoNumber

208                     If rsTrail.State = 1 Then rsTrail.Close
210                     rsTrail.Open "Select * from tblTrail where Username = '" & _
                                lbluser.Caption & "'"

                        'Updates new activity for trail activity.
212                     With rsTrail
214                         .AddNew
216                         !UserName = txtUser.Text
218                         !userlevel = lblUserlevel.Caption
220                         !Activity = "Add New Customer Record"
222                         !Time = lblTime.Caption
224                         !Date = DTPicker1.Value
226                         .Update
                        End With
                        
228                     txtSearch.Enabled = True

                        '  Unload Me
230                     frm_Customer.lbluser.Caption = MDIForm1.lblusername.Caption
232                     frm_Customer.txtUser.Text = MDIForm1.lblusername.Caption
234                     Me.Show
236                     btnClose_Click
                    End If
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnAdd_Click"

        Exit Sub

btnAdd_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.btnAdd_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnClose_Click()

        '<EhHeader>
        On Error GoTo btnClose_Click_Err

        TxtLog "Entered btnClose_Click"

        '</EhHeader>

100     txtSearch.Enabled = True

102     If btnClose.Caption = "&Close" Then
104         Unload Me
184         MDIForm1.Picture1.Visible = True
        Else
            ' txtCode.Text = ""
106         txtLastname = ""
108         txtFirstname = ""
110         txtMI.Text = ""
112         txtAddress.Text = ""
114         txtBalance.Text = ""
116         txtCLastname.Text = ""
118         txtCFirstname.Text = ""
120         cbCollector.Text = ""
122         cbCollector.Enabled = False
124         txtCFirstname.Enabled = False
126         txtCLastname.Enabled = False
128         txtCode.Enabled = False
130         txtLastname.Enabled = False
132         txtFirstname.Enabled = False
134         txtMI.Enabled = False
136         txtAddress.Enabled = False
138         txtRemarks.Text = ""
140         lblCode.Caption = ""
142         txtBalance.Enabled = False
144         txtRemarks.Enabled = False
146         btnClose.Caption = "&Close"
148         btnEdit.Caption = "&Edit"
150         btnEdit.Enabled = False
152         btnAdd.Caption = "&Add"
154         btnAdd.Enabled = True
156         btnReverse.Enabled = False
158         DataGrid1.Enabled = True
160         DataGrid1.Refresh
162         txtCode.Text = lblAuCode.Caption

164         If rsCustomer.State = 1 Then rsCustomer.Close
            'Sort records descending
166         rsCustomer.Open "Select * from tblCustomer Order By Code desc"
168         Set DataGrid1.DataSource = rsCustomer
170         DataGrid1.Width = Me.Width
172         DataGrid1.Columns(0).Width = 0
174         DataGrid1.Columns(1).Width = 550
176         DataGrid1.Columns(4).Width = 950
178         DataGrid1.Columns(5).Width = 4000
180         DataGrid1.Columns(7).Width = 2000
182         DataGrid1.Columns(8).Width = 1100

        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEdit_Click()

        '<EhHeader>
        On Error GoTo btnEdit_Click_Err

        TxtLog "Entered btnEdit_Click"

        '</EhHeader>

100     txtSearch.Enabled = False

102     If btnEdit.Caption = "&Edit" Then
104         btnEdit.Caption = "&Update"
106         txtLastname.Enabled = True
108         txtFirstname.Enabled = True
110         txtMI.Enabled = True
112         txtAddress.Enabled = True
114         cbCollector.Enabled = True
116         txtRemarks.Enabled = True
118         DataGrid1.Enabled = False
        Else
    
120         If rsCustomer.State = 1 Then rsCustomer.Close
122         rsCustomer.Open "Select * from tblCustomer where Code = " & txtCode.Text & ""

124         If txtLastname.Text = "" Or txtFirstname.Text = "" Or txtMI.Text = "" Or _
                    txtAddress.Text = "" Or cbCollector.Text = "" Or txtRemarks.Text = _
                    "" Then
126             MsgBox "All fields are required!", vbInformation, _
                        "Webplus Lending Corporation"
128             Call Customer
130             Set DataGrid1.DataSource = rsCustomer
               
132         ElseIf txtCode.Text = lblCode.Caption Then

134             If MsgBox("Are you sure you want to update this record?", vbQuestion + _
                        vbYesNo, "Webplus Lending Corporation") = vbYes Then

                    'Updates the desired customer record to database
136                 With rsCustomer
138                     !lastname = txtLastname.Text
140                     !firstname = txtFirstname.Text
142                     !MiddleInitial = txtMI.Text
144                     !Address = txtAddress.Text
146                     !Collector = txtCLastname.Text
148                     !Remarks = txtRemarks.Text
150                     !CollectorFirstname = txtCFirstname.Text
152                     !CollectorCode = cbCollector.Text
154                     .Update
                    End With

                    'Updating records of Customer on table Loans based on Customer Code
                    
156                 If rsLoan.State = 1 Then rsLoan.Close
158                 rsLoan.Open "Select * from tblLoan where Code = '" & _
                            rsCustomer!code & _
                            "' and (Status = 'Good' or Status = 'Full Paid') order by LoanID ASC"

160                 If rsLoan.RecordCount <> 0 Then
162                     rsLoan.MoveLast
                            
164                     rsLoan!Customer = txtLastname.Text
166                     rsLoan!firstname = txtFirstname.Text
168                     rsLoan!CollectorCode = cbCollector.Text
170                     rsLoan!Collector = txtCLastname.Text
172                     rsLoan!CollectorFname = txtCFirstname.Text
174                     rsLoan.Update
                                
176                     rsLoan.MoveNext
                        
                    End If
                    
                    'Updating records of Customer on table Payments based on Customer record
                    'needs improvement. it can be possible the only the date of encoding the it will edited that i
                    'it can also be possible that that payments the date of encoded will not be changed.
                    'only those payments that are encoded after.\
                    'so transfer tanan niya na payments? messy daw kaayo to.
            
178                 If rsTrail.State = 1 Then rsTrail.Close
180                 rsTrail.Open "Select * from tblTrail where Username = '" & _
                            lbluser.Caption & "'"

182                 With rsTrail
184                     .AddNew
186                     !UserName = lbluser.Caption
188                     !userlevel = lblUserlevel.Caption
190                     !Activity = "Edit Customer Record"
192                     !Time = lblTime.Caption
194                     !Date = lblDate.Caption
196                     .Update
                    End With

198                 txtSearch.Enabled = True
                    '  MsgBox " Record has Successfully Updated ", vbInformation, "Jajavi Lending Corporation"
200                 Me.Show
202                 btnClose_Click
                End If

            Else
            
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnEdit_Click"

        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.btnEdit_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnReverse_Click()

        '<EhHeader>
        On Error GoTo btnReverse_Click_Err

        TxtLog "Entered btnReverse_Click"

        '</EhHeader>

100     txtSearch.Enabled = False

102     If btnReverse.Caption = "&Edit" Then
104         btnReverse.Caption = "&Update"
106         txtLastname.Enabled = True
108         txtFirstname.Enabled = True
110         txtMI.Enabled = True
112         txtAddress.Enabled = True
114         cbCollector.Enabled = True
116         txtRemarks.Enabled = True
118         DataGrid1.Enabled = False
        Else
            'reverse save

120         If rsCustomer.State = 1 Then rsCustomer.Close
122         rsCustomer.Open "Select * from tblCustomer where Code =" & Val( _
                    txtCode.Text) & ""
        
            'check pud ani kung naa pa bay active loan ang customer?
            'possible ba kaha na ma zero ang iyaha balance bisag naa pa siyay active loan?
124         If txtLastname.Text = "" Or txtFirstname.Text = "" Or txtMI.Text = "" Or _
                    txtAddress.Text = "" Or cbCollector.Text = "" Or txtRemarks.Text = _
                    "" Then
126             MsgBox "All fields are required!", vbInformation, _
                        "Webplus Lending Corporation"
128             Call Customer
130             Set DataGrid1.DataSource = rsCustomer
132         ElseIf rsCustomer!Balance <> 0 Then
134             MsgBox _
                        "This customer has remaining balance.Unable to reverse this customer record", _
                        vbInformation, "Webplus Lending Corporation"
136             Call Customer
138             Set DataGrid1.DataSource = rsCustomer
140         ElseIf rsCustomer!Status = "Reversed" Then
142             MsgBox "This Customer is already reversed"
144             Call Customer
146             Set DataGrid1.DataSource = rsCustomer
            
148         ElseIf txtCode.Text = lblCode.Caption Then

150             If MsgBox("Are you sure to Reverse this record?", vbQuestion + vbYesNo, _
                        "Webplus Lending Corporation") = vbYes Then
                    
152                 With rsCustomer
                        '!lastname = txtLastName.Text
                        '!firstname = txtFirstName.Text
                        '!MiddleInitial = txtMI.Text
                        '!Address = txtAddress.Text
                        '!Collector = cbCollector.Text
154                     !Remarks = txtRemarks.Text
156                     !Status = "Reversed"
158                     .Update
                    End With

160                 txtSearch.Enabled = True
               
162                 frm_Customer.lbluser.Caption = MDIForm1.lblusername.Caption
164                 frm_Customer.txtUser.Text = MDIForm1.lblusername.Caption
166                 Me.Show
168                 btnClose_Click
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnReverse_Click"

        Exit Sub

btnReverse_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.btnReverse_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cbCollector_Click()

        '<EhHeader>
        On Error GoTo cbCollector_Click_Err

        TxtLog "Entered cbCollector_Click"

        '</EhHeader>

100     If rsCollector.State = 1 Then rsCollector.Close
102     rsCollector.Open "Select * from tblCollector where Code = " & cbCollector.Text _
                & " "

        '   Dim ColCode As String
104     If rsCollector.RecordCount <> 0 Then

106         txtCLastname.Text = rsCollector!lastname
108         txtCFirstname.Text = rsCollector!firstname
            '   lblCollectorCode.Caption = rsCollector!Code
            '   lblLName.Caption = rsCollector!LastName
            '   lblCollectorFname.Caption = rsCollector!FirstName
            '   cbCollector.Text = rsCollector!FirstName & ", " & rsCollector!LastName
        End If
               
110     If rsCollData.State = 1 Then rsCollData.Close
112     rsCollData.Open "Select * from tblColl_Data where Code = " & cbCollector.Text & _
                " order by DateEmployed DESC"

114     If rsCollData.RecordCount <> 0 Then
116         rsCollData.MoveFirst
118         txtCLastname.Text = rsCollData!lastname
120         txtCFirstname.Text = rsCollData!firstname
            'txtMI.Text = rsCollData!MI
            'dtEmployed.Value = rsCollData!DateEmployed
        Else
122         txtCLastname.Text = ""
124         txtCFirstname.Text = ""
        End If

        '<EhFooter>

        TxtLog "Exited cbCollector_Click"

        Exit Sub

cbCollector_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.cbCollector_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cbCollector_KeyDown(KeyCode As Integer, Shift As Integer)

        '<EhHeader>
        On Error GoTo cbCollector_KeyDown_Err

        TxtLog "Entered cbCollector_KeyDown"

        '</EhHeader>

        '  KeyCode = 0
100     If KeyCode = 13 Then
            'Call additems
        End If

        '<EhFooter>

        TxtLog "Exited cbCollector_KeyDown"

        Exit Sub

cbCollector_KeyDown_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.cbCollector_KeyDown", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cbCollector_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo cbCollector_KeyPress_Err

        TxtLog "Entered cbCollector_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
            ' Call additems
102         txtRemarks.SetFocus
104     ElseIf KeyAscii = 40 Then

            ' txtRemarks.SetFocus
106         If rsCollector.State = 1 Then rsCollector.Close
108         rsCollector.Open "Select * from tblCollector "
110         cbCollector.Clear

112         If rsCollector.RecordCount <> 0 Then

114             Do While Not rsCollector.EOF
116                 cbCollector.AddItem rsCollector!lastname
118                 rsCollector.MoveNext
                Loop

            End If

        Else
120         KeyAscii = 0
            
        End If

        '<EhFooter>

        TxtLog "Exited cbCollector_KeyPress"

        Exit Sub

cbCollector_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.cbCollector_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub
'Sub Search()
'If rsCustomer.State = 1 Then rsCustomer.Close
'rsCustomer.Open "Select * from tblCustomer where Code = '" & lblCode.Caption & "'"
    '         Set DataGrid1.DataSource = rsCustomer
'End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        TxtLog "Entered DataGrid1_Click"

        '</EhHeader>

        '       Call Customer
        '         lblCode.Caption = rsCustomer!Code
        '      If rsCustomer.State = 1 Then rsCustomer.Close
        '    rsCustomer.Open "Select * from tblCustomer where Code ='" & txtCode.Text & "'"
100     txtCode.Text = rsCustomer!code
            
        '      If rsCustomer.State = 1 Then rsCustomer.Close
        '     rsCustomer.Open "Select * from tblCustomer "
102     If rsCustomer.RecordCount = 0 Then
        
            'if rsCustomer.
        Else

            '    lblCode.Caption = rsCustomer!code
            '   rsCustomer.MoveNext
            ' rsCustomer.MoveNext
            '  If rsCustomer.State = 1 Then rsCustomer.Close
            ' rsCustomer.Open "Select * from tblCustomer where Code = '" & lblCode.Caption & "'"
            ' Set DataGrid1.DataSource = rsCustomer
            ' Call Search
            '  If rsCustomer.RecordCount = 1 Then
            '       If rsCustomer.State = 1 Then rsCustomer.Close
            '     rsCustomer.Open "Select * from tblCustomer where Code ='" & lblCode.Caption & "'"
104         With rsCustomer
106             txtCode.Text = !code
108             txtLastname.Text = !lastname
110             lblLastname.Caption = !lastname
112             txtFirstname.Text = !firstname
114             lblFirstname.Caption = !firstname
116             txtMI.Text = !MiddleInitial
118             txtAddress.Text = !Address
120             cbCollector.Text = !CollectorCode
122             lblCollector.Caption = !Collector
124             txtCLastname.Text = !Collector
126             txtCFirstname.Text = !CollectorFirstname
128             txtBalance.Text = !Balance
130             txtRemarks.Text = !Remarks
132             lblCode.Caption = !code

            End With

134         btnEdit.Enabled = True
136         btnAdd.Enabled = False
138         btnClose.Caption = "&Cancel"
            ' btnDelete.Enabled = True
140         btnReverse.Enabled = True
            
        End If

        ' End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_Click"

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.DataGrid1_Click", Erl

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
        ErrReport Err.Description, "LendingClient.frm_Customer.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

        'Me.WindowState = 2
        'Calls connection to database
100     Call connect
        'Calls recordset for Customer table
102     Call Customer
        'Calls recordset for Collector table
104     Call Collector
106     Call payment
108     Call Loan
110     Call CollCode
112     Call CollData
        
114     Call addItems
116     Me.Show
118     Me.SetFocus
120     Me.WindowState = vbMaximized
       
122     If rsCustomer.State = 1 Then rsCustomer.Close
        'Sort records descending
124     rsCustomer.Open "Select * from tblCustomer Order By Code desc"
126     Call autoNumber
128     Set DataGrid1.DataSource = rsCustomer
        'Adjusting the width of Fields on Datagird
130     DataGrid1.Width = Me.Width
132     DataGrid1.Columns(0).Width = 0
134     DataGrid1.Columns(1).Width = 550
136     DataGrid1.Columns(4).Width = 950
138     DataGrid1.Columns(5).Width = 4000
140     DataGrid1.Columns(7).Width = 2000
142     DataGrid1.Columns(8).Width = 1100
        '  Order By Code desc

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    lblTime.Caption = Time
    lblDate.Caption = Date
    
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtAddress_KeyPress_Err

        TxtLog "Entered txtAddress_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         cbCollector.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtAddress_KeyPress"

        Exit Sub

txtAddress_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.txtAddress_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtBalance_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtBalance_KeyPress_Err

        TxtLog "Entered txtBalance_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
        Else
        End If

        '<EhFooter>

        TxtLog "Exited txtBalance_KeyPress"

        Exit Sub

txtBalance_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.txtBalance_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtCode_KeyPress_Err

        TxtLog "Entered txtCode_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         txtLastname.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtCode_KeyPress"

        Exit Sub

txtCode_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.txtCode_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtFirstname_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtFirstname_KeyPress_Err

        TxtLog "Entered txtFirstname_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "ABCDEFGHIJKLMNOPQRSTUVWXYZÑabcdefghijklmnopqrstuvwxyzñ .-"

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
        ErrReport Err.Description, "LendingClient.frm_Customer.txtFirstname_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtLastname_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtLastname_KeyPress_Err

        TxtLog "Entered txtLastname_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "ABCDEFGHIJKLMNOPQRSTUVWXYZÑabcdefghijklmnopqrstuvwxyzñ- "

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then

104     ElseIf KeyAscii = 13 Then

106         txtFirstname.SetFocus
 
108     ElseIf KeyAscii = 8 Then
110         KeyAscii = 8

        Else
112         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtLastname_KeyPress"

        Exit Sub

txtLastname_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.txtLastname_KeyPress", Erl

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
        ErrReport Err.Description, "LendingClient.frm_Customer.txtMI_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtRemarks_KeyPress_Err

        TxtLog "Entered txtRemarks_KeyPress"

        '</EhHeader>

        '  Dim strvalid As String

        'strvalid = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"

        ' If InStr(1, strvalid, Chr(KeyAscii)) Then

100     If KeyAscii = 13 Then

102         btnAdd.SetFocus
 
            '  ElseIf KeyAscii = 8 Then
            '   KeyAscii = 8

            ' Else
            '    KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtRemarks_KeyPress"

        Exit Sub

txtRemarks_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.txtRemarks_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsCustomer.State = 1 Then rsCustomer.Close
102     rsCustomer.Open "Select * from tblCustomer where Lastname like '" & _
                txtSearch.Text & "%' or Firstname like '" & txtSearch.Text & _
                "%' or Code like '%" & txtSearch.Text & "%' Order by Code desc"

104     Set DataGrid1.DataSource = rsCustomer

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Customer.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

