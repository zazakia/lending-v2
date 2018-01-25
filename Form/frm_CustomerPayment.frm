VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_CustomerPayment 
   Caption         =   "frmPayment"
   ClientHeight    =   9375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   18960
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Customer Payment List"
      Height          =   9375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   18975
      Begin VB.TextBox txtSearch 
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   3855
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   8520
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   12726
         _Version        =   393216
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search here:"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frm_CustomerPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_CustomerPayment
'    Project    : Project1
'
'    Description: [hsdfsdfsdfsdfsdfsdfsdf]
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
        ErrReport Err.Description, "LendingClient.frm_CustomerPayment.btnClose_Click", Erl

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
        ErrReport Err.Description, _
                "LendingClient.frm_CustomerPayment.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call payment

104     Set DataGrid1.DataSource = rsPayment

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_CustomerPayment.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsPayment.State = 1 Then rsPayment.Close
102     rsPayment.Open "Select * from tblPayment where Customer like '" & _
                txtSearch.Text & "%' or ORnumber = '" & txtSearch.Text & _
                "%' or Code like '" & txtSearch.Text & "%'"
        
104     Set DataGrid1.DataSource = rsPayment

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CustomerPayment.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtSearch_KeyPress_Err

        TxtLog "Entered txtSearch_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz01234567890"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then

104     ElseIf KeyAscii = 13 Then
 
106     ElseIf KeyAscii = 8 Then
108         KeyAscii = 8

        Else

        End If

        '<EhFooter>

        TxtLog "Exited txtSearch_KeyPress"

        Exit Sub

txtSearch_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_CustomerPayment.txtSearch_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

