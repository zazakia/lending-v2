VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Delivery 
   Caption         =   "Delivery"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5265
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Charge Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtDelivery 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   3135
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   3480
         TabIndex        =   3
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1800
         TabIndex        =   2
         Top             =   3240
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3413
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
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Charge         :"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1380
      End
   End
End
Attribute VB_Name = "frm_Delivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Delivery
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
104         txtDelivery.Enabled = False
106         txtDelivery.Text = ""
108         btnEdit.Caption = "&Edit"
110         btnEdit.Enabled = False
112         btnClose.Caption = "&Close"
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Delivery.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnEdit_Click()

        '<EhHeader>
        On Error GoTo btnEdit_Click_Err

        TxtLog "Entered btnEdit_Click"

        '</EhHeader>

100     If btnEdit.Caption = "&Edit" Then

102         txtDelivery.Enabled = True
104         txtDelivery.SetFocus
106         btnEdit.Enabled = True
108         btnAdd.Enabled = True
110         btnClose.Caption = "&Cancel"
112         btnEdit.Caption = "&Update"
        Else

114         If txtDelivery.Text = "" Then
116             MsgBox "Please put the exact amount.", vbInformation
118             txtDelivery.SetFocus
            Else
    
120             If MsgBox(" Do you want to adjust the Charge rate?", vbQuestion + _
                        vbYesNo) = vbYes Then
  
122                 rsCharge!Amount = txtDelivery
124                 rsCharge.Update
    
126                 MsgBox "Delivery rate has been adjusted.", vbInformation
128                 Unload Me
130                 Me.Show
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnEdit_Click"

        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Delivery.btnEdit_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        TxtLog "Entered DataGrid1_Click"

        '</EhHeader>

100     If rsCharge.RecordCount = 0 Then

        Else
102         txtDelivery.Text = rsCharge!Amount
104         btnEdit.Enabled = True
106         btnAdd.Enabled = False
108         btnClose.Caption = "&Cancel"
        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_Click"

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Delivery.DataGrid1_Click", Erl

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
        ErrReport Err.Description, "LendingClient.frm_Delivery.DataGrid1_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call delivery
104     Call Charge
106     Set DataGrid1.DataSource = rsCharge

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Delivery.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtDelivery_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtDelivery_KeyPress_Err

        TxtLog "Entered txtDelivery_KeyPress"

        '</EhHeader>

        Dim strvalid As String

100     strvalid = "0987654321"

102     If InStr(1, strvalid, Chr$(KeyAscii)) Then
  
104     ElseIf KeyAscii = 8 Then
106         KeyAscii = 8
108     ElseIf KeyAscii = 13 Then
110         Call btnEdit_Click
    
        Else
112         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtDelivery_KeyPress"

        Exit Sub

txtDelivery_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Delivery.txtDelivery_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

