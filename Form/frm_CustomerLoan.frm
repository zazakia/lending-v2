VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_CustomerLoan 
   Caption         =   "frmCustomerLoan"
   ClientHeight    =   10005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   18960
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Customer Loan List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20175
      Begin VB.TextBox txtSearch 
         Height          =   495
         Left            =   8400
         TabIndex        =   3
         Top             =   480
         Width           =   4575
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   9120
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7815
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   19815
         _ExtentX        =   34951
         _ExtentY        =   13785
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
      Begin VB.Label Label1 
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
         Left            =   6720
         TabIndex        =   4
         Top             =   480
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frm_CustomerLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_CustomerLoan
'    Project    : Project1
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Private Sub btnClose_Click()

        '<EhHeader>
        On Error GoTo btnClose_Click_Err

        TxtLog "Entered btnClose_Click"

        '</EhHeader>

100     Unload Me

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_CustomerLoan.btnClose_Click", Erl

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
        ErrReport Err.Description, "LendingClient.frm_CustomerLoan.DataGrid1_KeyPress", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call Loan
104     Set DataGrid1.DataSource = rsLoan
106     DataGrid1.Columns(5).Caption = "Encoded Date"

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_CustomerLoan.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsLoan.State = 1 Then rsLoan.Close
102     rsLoan.Open "Select * from tblLoan where Customer like '" & txtSearch.Text & _
                "%' or Code like '%" & txtSearch.Text & "%'"
        
104     Set DataGrid1.DataSource = rsLoan

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_CustomerLoan.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtSearch_KeyPress_Err

        TxtLog "Entered txtSearch_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then

104     ElseIf KeyAscii = 13 Then
 
106     ElseIf KeyAscii = 8 Then
108         KeyAscii = 8

        Else
110         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtSearch_KeyPress"

        Exit Sub

txtSearch_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_CustomerLoan.txtSearch_KeyPress", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

