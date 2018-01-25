VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Servicefee 
   Caption         =   "frmServicefee"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox txtServicefee 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   735
         Left            =   2760
         TabIndex        =   1
         Top             =   3240
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1575
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2778
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
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2534.74
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Service Fee:"
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
         TabIndex        =   5
         Top             =   480
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frm_Servicefee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Servicefee
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

100     If btnClose.Caption = "&Close" Then
102         Unload Me
        Else
104         btnEdit.Caption = "&Edit"
106         btnEdit.Enabled = False
108         txtServicefee.Text = ""
110         btnClose.Caption = "&Close"
    
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Servicefee.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEdit_Click()

        '<EhHeader>
        On Error GoTo btnEdit_Click_Err

        TxtLog "Entered btnEdit_Click"

        '</EhHeader>

100     If btnEdit.Caption = "&Edit" Then
102         txtServicefee.Enabled = True
104         btnEdit.Caption = "&Update"
    
106     ElseIf btnEdit.Caption = "&Update" Then
    
108         rsServicefee!Servicefee = txtServicefee.Text
110         rsServicefee.Update
    
112         MsgBox "Service fee Successfully Updated"
    
        End If

        '<EhFooter>

        TxtLog "Exited btnEdit_Click"

        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Servicefee.btnEdit_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        TxtLog "Entered DataGrid1_Click"

        '</EhHeader>

100     If rsServicefee.RecordCount = 0 Then
        Else
102         txtServicefee.Text = rsServicefee!Servicefee

104         btnEdit.Enabled = True
106         btnClose.Caption = "&Cancel"

        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_Click"

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Servicefee.DataGrid1_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call Servicefee

104     Set DataGrid1.DataSource = rsServicefee

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Servicefee.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

