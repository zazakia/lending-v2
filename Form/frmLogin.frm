VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3795
   ClientLeft      =   7020
   ClientTop       =   3555
   ClientWidth     =   6210
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7080
      Top             =   4080
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   265289729
      CurrentDate     =   41886
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   1800
      Top             =   4080
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\New T\DB\JCashdb.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\New T\DB\JCashdb.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   6
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton btnLogin 
         Caption         =   "&Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtUsername 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1320
         TabIndex        =   10
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "LENDING CORPPORATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1155
         TabIndex        =   7
         Top             =   240
         Width           =   3900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password   :"
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
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username  :"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1530
      End
   End
   Begin VB.Label lblTime 
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   1680
      Width           =   375
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frmLogin
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
        ErrReport Err.Description, "LendingClient.frmLogin.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnLogin_Click()

        '<EhHeader>
        On Error GoTo btnLogin_Click_Err

        TxtLog "Entered btnLogin_Click"

        '</EhHeader>

        Dim userlevel As String

100     If txtUsername.Text = "" Or txtPassword.Text = "" Then
102         MsgBox "Username or password is empty. Please fill all required.", _
                    vbInformation, "Webplus Lending Corporation"
104         txtUsername.Text = ""
106         txtPassword.Text = ""
        Else

108         If rsUser.State = 1 Then rsUser.Close
110         rsUser.Open "select * from tblUser where Username = '" & Trim$( _
                    txtUsername.Text) & "' and Password = '" & Trim$(txtPassword.Text) _
                    & "'"

112         If rsUser.RecordCount <> 0 Then

114             userlevel = rsUser!userlevel
                
116             rsUser!Status = "Log-in"
118             rsUser.Update
120             MDIForm1.lblUserName.Caption = txtUsername.Text
        
122             frm_ChangePassword.lblpassword.Caption = rsUser!Password

124             If userlevel = "User" Then
126                 MDIForm1.users.Enabled = False
                    'frm_Collectors.btnDelete.Visible = False

128                 rsTrail.AddNew
130                 rsTrail!userlevel = userlevel
132                 rsTrail!UserName = Trim$(txtUsername.Text)
134                 rsTrail!Activity = "Logged In"
136                 rsTrail!Time = lblTime.Caption
138                 rsTrail!Date = DTPicker1.Value
140                 rsTrail!Status = "Log-in"
142                 rsTrail.Update
    
144                 MDIForm1.lblDate.Caption = DTPicker1.Value
146                 MDIForm1.lblUserlevel.Caption = userlevel
                    'frm_Collectors.btnDelete.Visible = False
                
148                 Unload Me
150                 Load MDIForm1
152                 MDIForm1.Show
                Else
154                 rsTrail.AddNew
156                 rsTrail!userlevel = userlevel
158                 rsTrail!UserName = Trim$(txtUsername.Text)
160                 rsTrail!Activity = "Logged In"
162                 rsTrail!Time = lblTime.Caption
164                 rsTrail!Date = DTPicker1.Value
166                 rsTrail!Status = "Log-in"
168                 rsTrail.Update
                         
170                 MDIForm1.lblDate.Caption = DTPicker1.Value
172                 MDIForm1.lblUserlevel.Caption = userlevel
174                 MDIForm1.lblUserName.Caption = Trim$(txtUsername.Text)
                   
                    'fm_Collectors.rbtnDelete.Visible = False
                
176                 Unload Me
178                 Load MDIForm1
180                 MDIForm1.Show

                End If

            Else
182             MsgBox "Invalid Username or password", vbInformation, _
                        "Webplus Lending Corporation"
184             txtUsername.Text = ""
186             txtPassword.Text = ""
188             txtUsername.SetFocus
            
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnLogin_Click"

        Exit Sub

btnLogin_Click_Err:
        ErrReport Err.Description, "LendingClient.frmLogin.btnLogin_Click", Erl

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
104     Call Trail

106     lblVersion = "Version: " & App.Major & "." & App.Minor & "." & App.Revision

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frmLogin.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    DTPicker1.Value = Date
    lblTime.Caption = Time
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtPassword_KeyPress_Err

        TxtLog "Entered txtPassword_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         btnLogin.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtPassword_KeyPress"

        Exit Sub

txtPassword_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frmLogin.txtPassword_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtUsername_KeyPress_Err

        TxtLog "Entered txtUsername_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         txtPassword.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtUsername_KeyPress"

        Exit Sub

txtUsername_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frmLogin.txtUsername_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

