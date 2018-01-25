VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form rep_cprForm 
   Caption         =   "rep_cprForm"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox cprSearch 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Width           =   5055
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   7335
      Left            =   720
      TabIndex        =   0
      Top             =   1560
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12938
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
   Begin VB.Label kita 
      Caption         =   "clickdetector"
      Height          =   255
      Left            =   9600
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   2160
   End
End
Attribute VB_Name = "rep_cprForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cprSearch_Change()

        '<EhHeader>
        On Error GoTo cprSearch_Change_Err

        TxtLog "Entered cprSearch_Change"

        '</EhHeader>

100     If rsCustomer.State = 1 Then rsCustomer.Close
102     rsCustomer.Open "Select * from tblCustomer where Lastname like '" & _
                cprSearch.Text & "%' or Firstname like '" & cprSearch.Text & _
                "%' or Code like '%" & cprSearch.Text & "%'"

104     Set DataGrid2.DataSource = rsCustomer

        '<EhFooter>

        TxtLog "Exited cprSearch_Change"

        Exit Sub

cprSearch_Change_Err:
        ErrReport Err.Description, "LendingClient.rep_cprForm.cprSearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid2_Click()

        '<EhHeader>
        On Error GoTo DataGrid2_Click_Err

        TxtLog "Entered DataGrid2_Click"

        '</EhHeader>
    
        'must pass chosen value to rep_cpr
100     With rsCustomer
102         kita.Caption = !code 'test
        
        End With

        '<EhFooter>

        TxtLog "Exited DataGrid2_Click"

        Exit Sub

DataGrid2_Click_Err:
        ErrReport Err.Description, "LendingClient.rep_cprForm.DataGrid2_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call Customer
104     Call Collector
106     Call payment
108     Call Loan
110     DataGrid2.Refresh
112     Set DataGrid2.DataSource = rsCustomer
114     DataGrid2.Refresh

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.rep_cprForm.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub kita_Change()

        '<EhHeader>
        On Error GoTo kita_Change_Err

        TxtLog "Entered kita_Change"

        '</EhHeader>

100     Unload rep_cpr
102     Load rep_cpr
104     rep_cpr.Show

        '<EhFooter>

        TxtLog "Exited kita_Change"

        Exit Sub

kita_Change_Err:
        ErrReport Err.Description, "LendingClient.rep_cprForm.kita_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

