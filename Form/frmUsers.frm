VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Users 
   Caption         =   "frmUsers"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   14040
      Top             =   1080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Password=vel123;Data Source=D:\lending cybergada\DB\JCashdb.mdb;Persist Security Info=True"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Password=vel123;Data Source=D:\lending cybergada\DB\JCashdb.mdb;Persist Security Info=True"
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
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   9840
      Width           =   10455
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   22
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton btnEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   21
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   16
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Personal details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   10455
      Begin VB.TextBox txtAddress 
         Enabled         =   0   'False
         Height          =   405
         Left            =   6360
         TabIndex        =   12
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtLastname 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtFirstname 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label13 
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
         Left            =   10200
         TabIndex        =   29
         Top             =   960
         Width           =   90
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
         Left            =   5520
         TabIndex        =   28
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Label11 
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
         Left            =   5520
         TabIndex        =   27
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Address      :"
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
         Left            =   6360
         TabIndex        =   11
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label5 
         Caption         =   "Last Name  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "First Name   :"
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
         TabIndex        =   7
         Top             =   480
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10455
      Begin VB.TextBox txtUserID 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   38
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   7680
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cbBranch 
         Height          =   315
         Left            =   7800
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   8520
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   265355265
         CurrentDate     =   41879
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   5160
         Top             =   720
      End
      Begin VB.ComboBox cbUserlevel 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmUsers.frx":0000
         Left            =   1800
         List            =   "frmUsers.frx":000A
         TabIndex        =   13
         Text            =   "User"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
         Enabled         =   0   'False
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtUsername 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label awe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID  : "
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
         Left            =   4080
         TabIndex        =   39
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Location :"
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
         Left            =   6480
         TabIndex        =   36
         Top             =   1440
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Branch    :"
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
         Left            =   6480
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label10 
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
         Left            =   3840
         TabIndex        =   26
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label Label9 
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
         Left            =   4800
         TabIndex        =   25
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label8 
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
         Left            =   4800
         TabIndex        =   24
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User  Level   :"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password    :"
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
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username   :"
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
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1350
      End
   End
   Begin VB.Frame Frame4 
      Height          =   6135
      Left            =   240
      TabIndex        =   17
      Top             =   3720
      Width           =   20055
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   3120
         TabIndex        =   19
         Top             =   120
         Width           =   4215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5415
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   19815
         _ExtentX        =   34951
         _ExtentY        =   9551
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Search here            :"
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
         Left            =   480
         TabIndex        =   20
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Label lblAuID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   11160
      TabIndex        =   40
      Top             =   840
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   8760
      TabIndex        =   37
      Top             =   5040
      Width           =   45
   End
   Begin VB.Label lblUserlevel 
      Caption         =   "Label14"
      Height          =   255
      Left            =   14400
      TabIndex        =   32
      Top             =   2400
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblTime 
      Height          =   255
      Left            =   13200
      TabIndex        =   31
      Top             =   2040
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblUser 
      Caption         =   "Label8"
      Height          =   255
      Left            =   12840
      TabIndex        =   23
      Top             =   2640
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "frm_Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frm_Users
'    Project    : Project1
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Sub addItems()

    '<EhHeader>
    On Error GoTo addItems_Err

    TxtLog "Entered addItems"

    '</EhHeader>

    '<EhFooter>

    TxtLog "Exited addItems"

    Exit Sub

addItems_Err:
    ErrReport Err.Description, "LendingClient.frm_Users.addItems", Erl

    Resume Next

    '</EhFooter>

End Sub

Sub autoNumber()

        '<EhHeader>
        On Error GoTo autoNumber_Err

        TxtLog "Entered autoNumber"

        '</EhHeader>

        Dim countUser As Integer

100     If rsUser.State = 1 Then rsUser.Close
102     rsUser.Open "Select * from tblUser Order By ID desc"
104     countUser = rsUser!ID + 1

106     If rsUser.RecordCount = 0 Then
108         txtUserID.Text = "1"
        Else
110         rsUser.MoveFirst
112         txtUserID.Text = "" & Format$(Right(countUser, 4), "")
114         lblAuID.Caption = txtUserID.Text
        End If

        '<EhFooter>

        TxtLog "Exited autoNumber"

        Exit Sub

autoNumber_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.autoNumber", Erl

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
104         btnClose.Caption = "&Cancel"
106         txtUsername.Enabled = True
108         txtPassword.Enabled = True
110         cbUserlevel.Enabled = True
112         txtFirstname.Enabled = True
114         txtLastname.Enabled = True
116         txtAddress.Enabled = True
118         txtUsername.SetFocus
120         DataGrid1.Enabled = False
        Else

122         If txtUsername.Text = "" Or txtPassword.Text = "" Or txtFirstname.Text = "" _
                    Or txtLastname.Text = "" Or txtAddress.Text = "" Then
124             MsgBox "All fields are required!", vbInformation, _
                        "Webplus Lending Corporation"
            Else

126             If rsUser.State = 1 Then rsUser.Close
128             rsUser.Open "Select * from tblUser where Username = '" & _
                        txtUsername.Text & "'"

130             If rsUser.RecordCount <> 0 Then
132                 MsgBox "User Already exist", vbInformation, _
                            "Webplus Lending Corporation"
134                 txtUsername.Text = ""
136                 txtUsername.SetFocus

                    Exit Sub

                Else

138                 If MsgBox("Are you sure you want to add new Record?", vbQuestion + _
                            vbYesNo, "Webplus Lending Corporation") = vbYes Then

140                     With rsUser
142                         .AddNew
144                         !UserID = txtUserID.Text
146                         !UserName = Trim$(txtUsername.Text)
148                         !Password = Trim$(txtPassword.Text)
150                         !userlevel = cbUserlevel.Text
152                         !lastname = Trim$(txtLastname.Text)
154                         !firstname = Trim$(txtFirstname.Text)
156                         !Address = Trim$(txtAddress.Text)
158                         !Status = "Log-out"
160                         .Update
                        End With
                
162                     If rsTrail.State = 1 Then rsTrail.Close
164                     rsTrail.Open "Select * from tblTrail "
                    
166                     With rsTrail
168                         .AddNew
170                         !UserName = lblUser.Caption
172                         !userlevel = lblUserlevel.Caption
174                         !Activity = "Add New User"
176                         !Time = lblTime.Caption
178                         !Date = DTPicker1.Value
180                         .Update
                        End With
    
182                     MsgBox "Record Successfully Added", vbInformation, _
                                "Webplus Lending Corporation"
184                     Call User
186                     Unload Me
188                     Me.Show
                    
                    End If
                
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnAdd_Click"

        Exit Sub

btnAdd_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.btnAdd_Click", Erl

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
        Else
104         btnAdd.Caption = "&Add"
106         txtUsername.Text = ""
108         txtPassword.Text = ""
110         cbUserlevel.Text = "User"
112         txtFirstname.Text = ""
114         txtLastname.Text = ""
116         txtAddress.Text = ""
118         txtUserID.Text = lblAuID.Caption
120         txtUsername.Enabled = False
122         txtPassword.Enabled = False
124         cbUserlevel.Enabled = False
126         txtFirstname.Enabled = False
128         txtLastname.Enabled = False
130         txtAddress.Enabled = False
    
132         btnClose.Caption = "&Close"
134         btnAdd.Enabled = True
136         btnEdit.Enabled = False
138         btnDelete.Enabled = False
140         DataGrid1.Enabled = True
    
        End If

        '<EhFooter>

        TxtLog "Exited btnClose_Click"

        Exit Sub

btnClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.btnClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub btnDelete_Click()

        '<EhHeader>
        On Error GoTo btnDelete_Click_Err

        TxtLog "Entered btnDelete_Click"

        '</EhHeader>

100     If rsUser.State = 1 Then rsUser.Close
102     rsUser.Open "Select * from tblUser where  UserID ='" & txtUserID.Text & "'"

104     If rsUser!Status = "Log-in" Then
106         MsgBox "This user is currently Logged in.You can't delete it.", _
                    vbInformation, "Webplus Lending Corp."
108         Call User
110         Set DataGrid1.DataSource = rsUser
112     ElseIf cbUserlevel.Text = "Admin" Then
114         MsgBox "Admin", vbInformation
        Else

116         If MsgBox("Are you sure you want to delete this user?", vbQuestion + _
                    vbYesNo) = vbYes Then
118             rsUser.Delete
120             rsUser.Update
122             MsgBox "Record successfully deleted", vbInformation
124             Unload Me
126             Me.Show
    
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnDelete_Click"

        Exit Sub

btnDelete_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.btnDelete_Click", Erl

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
    
104         txtUsername.Enabled = True
106         txtPassword.Enabled = True
108         txtLastname.Enabled = True
110         txtFirstname.Enabled = True
112         txtAddress.Enabled = True
114         cbUserlevel.Enabled = True
    
116         DataGrid1.Enabled = False
    
118     ElseIf btnEdit.Caption = "&Update" Then

120         If rsUser.State = 1 Then rsUser.Close
122         rsUser.Open "Select * from tblUser where UserID = '" & txtUserID.Text & "'"

124         If txtUsername.Text = "" Or txtPassword.Text = "" Or txtLastname.Text = "" _
                    Or txtFirstname.Text = "" Or txtAddress.Text = "" Then
126             MsgBox "All fiends are required! ", vbInformation
            Else

128             If MsgBox("Are you sure to update this record?", vbQuestion + vbYesNo, _
                        "J Lending Corporation") = vbYes Then
                  
                    Dim User      As String

                    Dim userlevel As String
                             
130                 With rsUser
132                     User = !UserName
134                     userlevel = !userlevel
136                     !UserName = txtUsername.Text
138                     !Password = txtPassword.Text
140                     !userlevel = cbUserlevel.Text
142                     !lastname = txtLastname.Text
144                     !firstname = txtFirstname.Text
146                     !Address = txtAddress.Text

148                     .Update
                    End With
                     
150                 rsUser.Close
                     
                    'edit other records update
                   
                    Dim ctr1 As Integer

152                 If rsCustomer.State = 1 Then rsCustomer.Close
154                 rsCustomer.Open "Select * from tblCustomer where User = '" & User & _
                            "'"

156                 If rsCustomer.RecordCount <> 0 Then
158                     rsCustomer.MoveFirst

160                     For ctr1 = 0 To rsCustomer.RecordCount
162                         rsCustomer.Close
164                         rsCustomer.Open "Select * from tblCustomer where User = '" _
                                    & User & "'"

166                         If rsCustomer.RecordCount <> 0 Then
168                             rsCustomer!User = txtUsername.Text
170                             rsCustomer.Update
                            End If

172                     Next ctr1

                    End If

174                 rsCustomer.Close
                   
                    'expense
                 
                    Dim ctr2 As Integer
                 
176                 If rsExpense.State = 1 Then rsExpense.Close
178                 rsExpense.Open "Select * from tblExpense"

180                 If rsExpense.RecordCount <> 0 Then
182                     rsExpense.MoveFirst
                 
184                     For ctr2 = 0 To rsExpense.RecordCount
186                         rsExpense.Close
188                         rsExpense.Open "Select * from tblExpense where User = '" & _
                                    User & "'"

190                         If rsExpense.RecordCount <> 0 Then
192                             rsExpense!User = txtUsername.Text
194                             rsExpense.Update
                            End If

196                     Next ctr2
                 
                    End If

198                 rsExpense.Close
                 
                    'expense end
                
                    'trail
                 
                    Dim ctr3 As Integer

200                 If rsTrail.State = 1 Then rsTrail.Close
202                 rsTrail.Open "Select * from tblTrail where Username = '" & User & "'"

204                 If rsTrail.RecordCount <> 0 Then
206                     rsTrail.MoveFirst
                
208                     For ctr3 = 0 To rsTrail.RecordCount
210                         rsTrail.Close
212                         rsTrail.Open "Select * from tblTrail where Username = '" & _
                                    User & "'"

214                         If rsTrail.RecordCount <> 0 Then
216                             rsTrail!UserName = txtUsername.Text
218                             rsTrail!userlevel = userlevel
220                             rsTrail.Update
                            End If

222                     Next ctr3

                    End If

224                 rsTrail.Close
                    'trail end
                   
                    'loan
                    Dim ctr4 As Integer

226                 If rsLoan.State = 1 Then rsLoan.Close
228                 rsLoan.Open "Select * from tblLoan where User = '" & User & "'"

230                 If rsLoan.RecordCount <> 0 Then
232                     rsLoan.MoveFirst
                 
234                     For ctr4 = 0 To rsLoan.RecordCount
236                         rsLoan.Close
238                         rsLoan.Open "Select * from tblLoan where User = '" & User & _
                                    "'"

240                         If rsLoan.RecordCount <> 0 Then
242                             rsLoan!User = txtUsername.Text
244                             rsLoan.Update
                            End If

246                     Next ctr4

                    End If

248                 rsLoan.Close
                    'loan end
                
                    'payment
                    Dim ctr5 As Integer

250                 If rsPayment.State = 1 Then rsPayment.Close
252                 rsPayment.Open "Select * from tblPayment where User = '" & User & "'"

254                 If rsPayment.RecordCount <> 0 Then
256                     rsPayment.MoveFirst
                 
258                     For ctr5 = 0 To rsPayment.RecordCount
260                         rsPayment.Close
262                         rsPayment.Open "Select * from tblPayment where User = '" & _
                                    User & "'"

264                         If rsPayment.RecordCount <> 0 Then
266                             rsPayment!User = txtUsername.Text
268                             rsPayment.Update
                            End If

270                     Next ctr5

                    End If

272                 rsPayment.Close
                    'payment end
                
                    'breakdown
                    Dim ctr6 As Integer

274                 If rsBreak.State = 1 Then rsBreak.Close
276                 rsBreak.Open "Select * from tblBreakdown where User = '" & User & "'"

278                 If rsBreak.RecordCount <> 0 Then
280                     rsBreak.MoveFirst
                 
282                     For ctr6 = 0 To rsBreak.RecordCount
284                         rsBreak.Close
286                         rsBreak.Open "Select * from tblBreakdown where User = '" & _
                                    User & "'"

288                         If rsBreak.RecordCount <> 0 Then
290                             rsBreak!User = txtUsername.Text
292                             rsBreak.Update
                            End If

294                     Next ctr6

                    End If

296                 rsBreak.Close
                    'breakdown end
                
                    'Cash on Hand
                    Dim ctr7 As Integer

298                 If rsCashOnHand.State = 1 Then rsCashOnHand.Close
300                 rsCashOnHand.Open "Select * from tblCashOnHand where User = '" & _
                            User & "'"

302                 If rsCashOnHand.RecordCount <> 0 Then
304                     rsCashOnHand.MoveFirst
                 
306                     For ctr7 = 0 To rsCashOnHand.RecordCount
308                         rsCashOnHand.Close
310                         rsCashOnHand.Open _
                                    "Select * from tblCashOnHand where User = '" & User _
                                    & "'"

312                         If rsCashOnHand.RecordCount <> 0 Then
314                             rsCashOnHand!User = txtUsername.Text
316                             rsCashOnHand.Update
                            End If

318                     Next ctr7

                    End If

320                 rsCashOnHand.Close
                    'Cash on Hand end
                 
                    'cashonbank
                    Dim ctr8 As Integer

322                 If rsCashOnBank.State = 1 Then rsCashOnBank.Close
324                 rsCashOnBank.Open "Select * from tblCashOnBank where User = '" & _
                            User & "'"

326                 If rsCashOnBank.RecordCount <> 0 Then
328                     rsCashOnBank.MoveFirst
                 
330                     For ctr8 = 0 To rsCashOnBank.RecordCount
332                         rsCashOnBank.Close
334                         rsCashOnBank.Open _
                                    "Select * from tblCashOnBank where User = '" & User _
                                    & "'"

336                         If rsCashOnBank.RecordCount <> 0 Then
338                             rsCashOnBank!User = txtUsername.Text
340                             rsCashOnBank.Update
                            End If

342                     Next ctr8

                    End If

344                 rsCashOnBank.Close
                    'cashonhand
                 
                    'edit other records update end
                
346                 If rsTrail.State = 1 Then rsTrail.Close
348                 rsTrail.Open "Select * from tblTrail "
                    
350                 With rsTrail
352                     .AddNew
354                     !UserName = lblUser.Caption
356                     !userlevel = lblUserlevel.Caption
358                     !Activity = "Add New User"
360                     !Time = lblTime.Caption
362                     !Date = DTPicker1.Value
364                     .Update
                    End With
                    
366                 MsgBox "User successfully Updated", vbInformation, "Webplus Lending"
368                 Unload Me
370                 Me.Show
                End If
            End If
        End If

        '<EhFooter>

        TxtLog "Exited btnEdit_Click"

        Exit Sub

btnEdit_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.btnEdit_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cbBranch_Click()

        '<EhHeader>
        On Error GoTo cbBranch_Click_Err

        TxtLog "Entered cbBranch_Click"

        '</EhHeader>

100     If rsBranch.State = 1 Then rsBranch.Close
102     rsBranch.Open "Select * from tblBranch where BranchName = '" & cbBranch.Text & "'"
104     txtLocation.Text = rsBranch!BranchLocation

        '<EhFooter>

        TxtLog "Exited cbBranch_Click"

        Exit Sub

cbBranch_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.cbBranch_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cbUserlevel_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo cbUserlevel_KeyPress_Err

        TxtLog "Entered cbUserlevel_KeyPress"

        '</EhHeader>

100     KeyAscii = 0

        '<EhFooter>

        TxtLog "Exited cbUserlevel_KeyPress"

        Exit Sub

cbUserlevel_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.cbUserlevel_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        TxtLog "Entered DataGrid1_Click"

        '</EhHeader>

100     If rsUser.RecordCount = 0 Then
        Else
102         btnAdd.Enabled = False
104         btnClose.Caption = "&Cancel"
106         btnEdit.Enabled = True
108         btnDelete.Enabled = True
    
110         With rsUser
112             txtUsername.Text = !UserName
114             txtPassword.Text = !Password
116             cbUserlevel.Text = !userlevel
118             txtLastname.Text = !lastname
120             txtFirstname.Text = !firstname
122             txtAddress.Text = !Address
124             txtUserID.Text = !UserID
126             lblID.Caption = !ID
                
            End With
    
        End If

        '<EhFooter>

        TxtLog "Exited DataGrid1_Click"

        Exit Sub

DataGrid1_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.DataGrid1_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DataGrid1_KeyPress_Err

        TxtLog "Entered DataGrid1_KeyPress"

        '</EhHeader>

100     KeyAscii = 13

        '<EhFooter>

        TxtLog "Exited DataGrid1_KeyPress"

        Exit Sub

DataGrid1_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.DataGrid1_KeyPress", Erl

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
104     Call Branch
        'Set rsExpense = Nothing
106     Call Expense
        'Set rsCustomer = Nothing
108     Call Customer
        'Set rsTrail = Nothing
110     Call Trail
112     Call Loan
114     Call payment
116     Call Break
118     Call Cashonhand
120     Call cashonbank

122     Call addItems

124     Call autoNumber

126     If rsUser.State = 1 Then rsUser.Close
        'Sort records descending
128     rsUser.Open "Select * from tblUser Order By ID desc"
130     Set DataGrid1.DataSource = rsUser
132     DataGrid1.Width = Me.Width
134     DataGrid1.Columns(0).Width = 500
136     DataGrid1.Columns(1).Width = 0
138     DataGrid1.Columns(3).Visible = False

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Timer1_Timer()

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    DTPicker1.Value = Date
    lblTime.Caption = Time
    Timer1.Enabled = False
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtAddress_KeyPress_Err

        TxtLog "Entered txtAddress_KeyPress"

        '</EhHeader>

100     If KeyAscii = 13 Then
102         btnAdd.SetFocus
        End If

        '<EhFooter>

        TxtLog "Exited txtAddress_KeyPress"

        Exit Sub

txtAddress_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.txtAddress_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If rsUser.State = 1 Then rsUser.Close
102     rsUser.Open "Select * from tblUser where Username like '" & txtSearch.Text & _
                "%' order by UserID desc"
104     Set DataGrid1.DataSource = rsUser

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Users.txtsearch_Change", Erl

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
        ErrReport Err.Description, "LendingClient.frm_Users.txtUsername_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

