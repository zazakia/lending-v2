VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Trail 
   Caption         =   "frmTrail"
   ClientHeight    =   10425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   18960
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Trail User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   11175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   18975
      Begin VB.TextBox txtText1 
         Height          =   405
         Left            =   10920
         TabIndex        =   5
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtsearch 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   1200
         Width           =   4695
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "Close"
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   10320
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid dtTrail 
         Height          =   8295
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   14631
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
      Begin VB.Label lblSEARCHFOROO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH FOR ACTIVITY DATE"
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
         Left            =   10920
         TabIndex        =   7
         Top             =   840
         Width           =   3435
      End
      Begin VB.Label lblSEARCHFOR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH FOR USERNAME OR ACTIVITY"
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
         Left            =   3240
         TabIndex        =   6
         Top             =   840
         Width           =   4590
      End
      Begin VB.Label lblSearchHere 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Here: "
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
         Left            =   1440
         TabIndex        =   4
         Top             =   1200
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frm_Trail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       btnClose_Click
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       GENUINE
' Date-Time  :       8/27/2014-8:37:05 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
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
        ErrReport Err.Description, "LendingClient.frm_Trail.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub dtTrail_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo dtTrail_KeyPress_Err

        TxtLog "Entered dtTrail_KeyPress"

        '</EhHeader>

100     KeyAscii = 0

        '<EhFooter>

        TxtLog "Exited dtTrail_KeyPress"

        Exit Sub

dtTrail_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Trail.dtTrail_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       Project1
' Procedure  :       Form_Load
' Description:       [type_description_here]
' Created by :       Project Administrator
' Machine    :       GENUINE
' Date-Time  :       8/27/2014-8:37:05 AM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call Trail

104     If rsTrail.State = 1 Then rsTrail.Close
106     rsTrail.Open "Select * from tblTrail where Activity like  '" & txtSearch.Text & _
                "%' or Username like '" & txtSearch.Text & "%' Order by ID desc"
108     Set dtTrail.DataSource = rsTrail

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Trail.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsTrail.State = 1 Then rsTrail.Close
102     rsTrail.Open "Select * from tblTrail where Activity like  '" & txtSearch.Text & _
                "%' or Username like '" & txtSearch.Text & "%' Order by ID desc"
104     Set dtTrail.DataSource = rsTrail

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Trail.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtText1_Change()

        '<EhHeader>
        On Error GoTo txtText1_Change_Err

        TxtLog "Entered txtText1_Change"

        '</EhHeader>

100     If rsTrail.State = 1 Then rsTrail.Close
102     rsTrail.Open "Select * from tblTrail where Date like  '" & txtText1.Text & _
                "%' Order by ID desc"
104     Set dtTrail.DataSource = rsTrail

        '<EhFooter>

        TxtLog "Exited txtText1_Change"

        Exit Sub

txtText1_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Trail.txtText1_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

