VERSION 5.00
Begin VB.Form frm_ChangePassword 
   Caption         =   "frmChangePassword"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         Height          =   615
         Left            =   3120
         TabIndex        =   8
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton btnChangepassword 
         Caption         =   "&Change"
         Height          =   615
         Left            =   480
         TabIndex        =   7
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txtConfirmpassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txtNewpassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtOldpassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password        :"
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
         Width           =   2460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "New Password              :"
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
         Top             =   1560
         Width           =   2460
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Old Password                :"
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
         Width           =   2475
      End
   End
   Begin VB.Label lblpassword 
      Caption         =   "lblpassword"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "frm_ChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_ChangePassword
'    Project    : Project1
'
'    Description: [This module will change the password of the user]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Private Sub btnChangepassword_Click()

        '<EhHeader>
        On Error GoTo btnChangepassword_Click_Err

        TxtLog "Entered btnChangepassword_Click"

        '</EhHeader>

100     If txtOldpassword.Text = "" Then
102         MsgBox "Please type your old password.", vbInformation
104         txtOldpassword.Text = ""
106         txtOldpassword.SetFocus
108     ElseIf txtNewpassword.Text = "" Then
110         MsgBox "Please type your new password.", vbInformation
112         txtNewpassword.Text = ""
114         txtNewpassword.SetFocus
116     ElseIf txtConfirmpassword.Text = "" Then
118         MsgBox "Please confirm your password.", vbInformation
120         txtConfirmpassword.Text = ""
122         txtConfirmpassword.SetFocus
124     ElseIf txtOldpassword.Text <> lblpassword.Caption Then
126         MsgBox "Old password did not match.", vbInformation
128         txtOldpassword.Text = ""
130         txtOldpassword.SetFocus
132     ElseIf txtNewpassword.Text <> txtConfirmpassword.Text Then
134         MsgBox "Password did not match.", vbInformation
136         txtNewpassword.Text = ""
138         txtConfirmpassword.Text = ""
140         txtNewpassword.SetFocus
        Else

142         If MsgBox("Are you sure you want to change you password?", vbQuestion + _
                    vbYesNo, "Webplus Lending") = vbYes Then

144             If rsUser.State = 1 Then rsUser.Close
146             rsUser.Open "Select * from tblUser where Password = '" & _
                        lblpassword.Caption & "'"

148             If rsUser.RecordCount <> 0 Then
        
150                 rsUser!Password = Trim$(txtNewpassword.Text)
152                 rsUser.Update
154                 MsgBox "Your password is successfully changed.", vbInformation
156                 Unload Me
158                 MDIForm1.Hide
160                 Load frmLogin
162                 frmLogin.Show
            
                End If
            End If
    
        End If

        '<EhFooter>

        TxtLog "Exited btnChangepassword_Click"

        Exit Sub

btnChangepassword_Click_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_ChangePassword.btnChangepassword_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

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
        ErrReport Err.Description, "LendingClient.frm_ChangePassword.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

100     Call connect
102     Call User

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_ChangePassword.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtConfirmpassword_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtConfirmpassword_KeyPress_Err

        TxtLog "Entered txtConfirmpassword_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         btnChangepassword.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtConfirmpassword_KeyPress"

        Exit Sub

txtConfirmpassword_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_ChangePassword.txtConfirmpassword_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtNewpassword_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtNewpassword_KeyPress_Err

        TxtLog "Entered txtNewpassword_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         txtConfirmpassword.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtNewpassword_KeyPress"

        Exit Sub

txtNewpassword_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_ChangePassword.txtNewpassword_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtOldpassword_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtOldpassword_KeyPress_Err

        TxtLog "Entered txtOldpassword_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         txtNewpassword.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtOldpassword_KeyPress"

        Exit Sub

txtOldpassword_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_ChangePassword.txtOldpassword_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

