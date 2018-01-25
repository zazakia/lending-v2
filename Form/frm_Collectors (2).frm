VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_Collectors 
   Caption         =   " Collectors"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   600
      Left            =   9120
      TabIndex        =   35
      Top             =   4320
      Width           =   990
   End
   Begin MSComCtl2.DTPicker dtSwitch 
      Height          =   375
      Left            =   7440
      TabIndex        =   30
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   118882305
      CurrentDate     =   42055
   End
   Begin VB.Frame FraCurrColl 
      Caption         =   "Current Collectors"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   10560
      TabIndex        =   28
      Top             =   1200
      Width           =   9015
      Begin VB.Label lblCurrColl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   29
         Top             =   600
         Width           =   75
      End
   End
   Begin VB.ComboBox cmbCollCode2 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdSwitchArea 
      Caption         =   "Switch Area"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5400
      TabIndex        =   26
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3720
      TabIndex        =   25
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7320
      TabIndex        =   21
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox cmbCollCode 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   480
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DGCollectors 
      Height          =   4695
      Left            =   10560
      TabIndex        =   19
      Top             =   5760
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8281
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11640
      TabIndex        =   18
      Text            =   "Search"
      Top             =   5160
      Width           =   5535
   End
   Begin VB.CommandButton cmdReplaceCollector 
      Caption         =   "Replace Collector"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1440
      TabIndex        =   15
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      TabIndex        =   14
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdAddCollector 
      Caption         =   "Add Collector Area"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3360
      TabIndex        =   13
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame framePayment_made 
      Caption         =   "Payment Made"
      Height          =   2295
      Left            =   240
      TabIndex        =   31
      Top             =   960
      Width           =   9855
      Begin VB.Label label_orandname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OR and name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   34
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   240
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   750
      End
      Begin VB.Label LabelCollectInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Collect info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   240
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.Frame FraCollector 
      Caption         =   "Collector"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   9855
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   4
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtStatus 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7200
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtEmployed 
         Height          =   375
         Left            =   7200
         TabIndex        =   5
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   119144449
         CurrentDate     =   42045
      End
      Begin VB.TextBox txtLastName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtMI 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   3
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtFirstName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblAddress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5760
         TabIndex        =   11
         Top             =   1200
         Width           =   1410
      End
      Begin VB.Label lblDateEmployed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Employed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5760
         TabIndex        =   10
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label lblMI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M.I."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   930
      End
      Begin VB.Label lblLastName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label lblFirstName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.Label lblSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10800
      TabIndex        =   24
      Top             =   5280
      Width           =   675
   End
   Begin VB.Label lblCollectorID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CollectorID"
      Height          =   195
      Left            =   1200
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   11160
      TabIndex        =   17
      Top             =   480
      Width           =   75
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10560
      TabIndex        =   16
      Top             =   480
      Width           =   525
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code :            :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   1620
   End
End
Attribute VB_Name = "frm_Collectors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Brayan 03-16-2015 12:00 AM
'mo ni ang function kung magdelete ug collector kinahanglan mabalik ang information sa old collector nga gepuliha
Public Sub RestCollectorInformation(x As Integer)

        '<EhHeader>
        On Error GoTo RestCollectorInformation_Err

        TxtLog "Entered RestCollectorInformation"

        '</EhHeader>

100     If rsCollData.State = 1 Then rsCollData.Close
102     rsCollData.Open "Select * From tblColl_Data Where Code = '" & x & _
                "' ORDER BY DateEmployed DESC"
    
104     If rsCollector.State = 1 Then rsCollector.Close
106     rsCollector.Open "Select * From tblCollector Where Code= '" & x & "'"

108     With rsCollector
110         !firstname = rsCollData!firstname
112         !lastname = rsCollData!lastname
114         !MiddleInitial = rsCollData!MI
116         !Address = rsCollData!Adress
118         !DateEmployed = rsCollData!DateEmployed
120         !collection = rsCollData!collection
122         !YTDCollection = rsCollData!YTDCollection
124         .Update
        End With

        '<EhFooter>

        TxtLog "Exited RestCollectorInformation"

        Exit Sub

RestCollectorInformation_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_Collectors.RestCollectorInformation", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub setgrid()

        '<EhHeader>
        On Error GoTo setgrid_Err

        TxtLog "Entered setgrid"

        '</EhHeader>

100     txtSearch.Text = ""

102     If rsCollData2.State = 1 Then rsCollData2.Close
104     rsCollData2.Open "Select * from tblColl_Data order by ID desc"
        'rsPayment.Open "Select * from tblPayment"
    
106     Set DGCollectors.DataSource = rsCollData2
108     DGCollectors.Columns(0).Width = 0
110     DGCollectors.Columns(1).Width = 700
112     DGCollectors.Columns(4).Width = 400
114     DGCollectors.Columns(5).Width = 3200
116     DGCollectors.Columns(7).Width = 0
118     DGCollectors.Refresh

        '<EhFooter>

        TxtLog "Exited setgrid"

        Exit Sub

setgrid_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.setgrid", Erl

        Resume Next

        '</EhFooter>

End Sub

'Jurey Tavera 03-16-2015 12:00 AM
'Mao ni ang funcion sa pagswitch sa tblCollector
Public Sub SwitchInfoTblCollector(code, firstname, lastname, MI, Adress, DEmp, coll, YTD)

        '<EhHeader>
        On Error GoTo SwitchInfoTblCollector_Err

        TxtLog "Entered SwitchInfoTblCollector"

        '</EhHeader>

100     If rsCollector.State = 1 Then rsCollector.Close
102     rsCollector.Open "Select * FROM tblCollector Where Code = '" & code & _
                "' ORDER BY DateEmployed DESC"

104     With rsCollector
106         !firstname = firstname
108         !lastname = lastname
110         !MiddleInitial = MI
112         !Address = Adress
114         !DateEmployed = DEmp
116         !collection = coll
118         !YTDCollection = YTD
120         .Update
        End With

        '<EhFooter>

        TxtLog "Exited SwitchInfoTblCollector"

        Exit Sub

SwitchInfoTblCollector_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_Collectors.SwitchInfoTblCollector", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub clear_data()

        '<EhHeader>
        On Error GoTo clear_data_Err

        TxtLog "Entered clear_data"

        '</EhHeader>

100     txtLastname.Text = ""
102     txtFirstName.Text = ""
104     txtMI.Text = ""
106     txtAddress.Text = ""
108     dtEmployed.Value = DateTime.Date
110     txtStatus.Text = ""

        'DtReplaced.Value = DateTime.Date

        '<EhFooter>

        TxtLog "Exited clear_data"

        Exit Sub

clear_data_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.clear_data", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub get_current_collectors()

        '<EhHeader>
        On Error GoTo get_current_collectors_Err

        TxtLog "Entered get_current_collectors"

        '</EhHeader>

100     lblCurrColl.Caption = "" '"heheh" & vbNewLine & "heheh"
102     FraCurrColl.Height = 1215
    
104     If rsCollCode.State = 1 Then rsCollCode.Close
106     rsCollCode.Open "Select * from tblColl_Code"
108     cmbCollCode.Clear
        
110     If rsCollCode.RecordCount <> 0 Then
                
112         Do While Not rsCollCode.EOF
        
114             If rsCollData2.State = 1 Then rsCollData2.Close
116             rsCollData2.Open "Select * from tblColl_Data where Code = " & _
                        rsCollCode!code & " order by DateEmployed desc"

118             If rsCollData2.RecordCount <> 0 Then
120                 lblCurrColl.Caption = lblCurrColl.Caption & vbNewLine & _
                            rsCollData2!code & ": " & rsCollData2!lastname & ", " & _
                            rsCollData2!firstname & " " & rsCollData2!MI & "."
122                 FraCurrColl.Height = FraCurrColl.Height + 285
                End If

124             rsCollCode.MoveNext
            Loop
                
        End If

        '<EhFooter>

        TxtLog "Exited get_current_collectors"

        Exit Sub

get_current_collectors_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_Collectors.get_current_collectors", Erl

        Resume Next

        '</EhFooter>

End Sub

'Benjamin Sumilhig 03-16-2015 7:50 PM
'Get the complete name of the customer
Sub get_name(ByRef ORnum As String, _
             ByRef customer_code As String, _
             ByRef paymentsMade As Double)

        '<EhHeader>
        On Error GoTo get_name_Err

        TxtLog "Entered get_name"

        '</EhHeader>

        Dim complete_name As String

        'this hold the value of the customer code that pass by the procedure
    
100     If rsCustomer.State = 1 Then rsCustomer.Close
102     rsCustomer.Open "SELECT * FROM tblCustomer WHERE Code = " & customer_code & ""
        ' this holds the complete name of the customer that uses to return a value
104     complete_name = ORnum & ": " & rsCustomer!lastname & ", " & _
                rsCustomer!firstname & " " & rsCustomer!MiddleInitial & _
                ".    |    Payments Made : " & paymentsMade
106     label_orandname.Caption = label_orandname.Caption & vbNewLine & complete_name

        '<EhFooter>

        TxtLog "Exited get_name"

        Exit Sub

get_name_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.get_name", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub load_coll_code()

        '<EhHeader>
        On Error GoTo load_coll_code_Err

        TxtLog "Entered load_coll_code"

        '</EhHeader>

100     If rsCollCode.State = 1 Then rsCollCode.Close
102     rsCollCode.Open "Select * from tblColl_Code"
104     cmbCollCode.Clear

        'Benjamin Sumilhig 03-12-2015 2-15-2015
106     If rsCollCode.RecordCount <> 0 Then
                
108         Do While Not rsCollCode.EOF
110             cmbCollCode.AddItem rsCollCode!code
112             rsCollCode.MoveNext
            Loop

        Else
114         cmbCollCode.AddItem 0
        End If

116     cmbCollCode.ListIndex = 0

        '<EhFooter>

        TxtLog "Exited load_coll_code"

        Exit Sub

load_coll_code_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.load_coll_code", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub load_coll_code2()

        '<EhHeader>
        On Error GoTo load_coll_code2_Err

        TxtLog "Entered load_coll_code2"

        '</EhHeader>

100     If rsCollCode.State = 1 Then rsCollCode.Close
102     rsCollCode.Open "Select * from tblColl_Code"
104     cmbCollCode2.Clear
        
106     If rsCollCode.RecordCount <> 0 Then
                
108         Do While Not rsCollCode.EOF
110             cmbCollCode2.AddItem rsCollCode!code
112             rsCollCode.MoveNext
            Loop
                
        End If
    
114     cmbCollCode2.ListIndex = rsCollCode.RecordCount - 1

        '<EhFooter>

        TxtLog "Exited load_coll_code2"

        Exit Sub

load_coll_code2_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.load_coll_code2", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub set_dp()

        '<EhHeader>
        On Error GoTo set_dp_Err

        TxtLog "Entered set_dp"

        '</EhHeader>

100     dtEmployed.Value = DateTime.Date

        'DtReplaced.Value = DateTime.Date

        '<EhFooter>

        TxtLog "Exited set_dp"

        Exit Sub

set_dp_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.set_dp", Erl

        Resume Next

        '</EhFooter>

End Sub

Sub txtSearch_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtSearch_KeyPress_Err

        TxtLog "Entered txtSearch_KeyPress"

        '</EhHeader>

        Dim strvalid As String

        'Benjamin sumilhig
        'Nag add me ug special charater Para maka type siya ug Ò ug —

100     strvalid = _
                "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZÒ—0987654321/!@#$%^&*()_+=-'\.,|[]{} "

102     If (InStr(1, strvalid, Chr$(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 13) = 0 Then
104         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtSearch_KeyPress"

        Exit Sub

txtSearch_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.txtSearch_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmbCollCode2_Click()

        '<EhHeader>
        On Error GoTo cmbCollCode2_Click_Err

        TxtLog "Entered cmbCollCode2_Click"

        '</EhHeader>

100     If cmbCollCode.Text = cmbCollCode2.Text Then
            'MsgBox (cmbCollCode.ListIndex)
102         MsgBox ("Cannot switch with the same area")
        
            Dim cmbCount As Integer
        
104         For cmbCount = 0 To cmbCollCode.ListCount - 1
106             cmbCollCode2.ListIndex = cmbCount

108             If cmbCollCode.Text <> cmbCollCode2.Text Then

                    Exit For

                End If

110         Next cmbCount

        End If
    
112     If rsCollData.State = 1 Then rsCollData.Close
114     rsCollData.Open "Select * from tblColl_Data where Code = " & cmbCollCode2.Text _
                & " order by DateEmployed DESC"

116     If rsCollData.RecordCount <> 0 Then
    
118         If cmdAddCollector.Caption <> "Confirm Switch" Then
        
120             cmdAddCollector.Enabled = False
122             cmdEdit.Enabled = True
124             cmdReplaceCollector.Enabled = True
126             cmdDelete.Enabled = True
            
128             txtFirstName.Enabled = False
130             txtLastname.Enabled = False
132             txtMI.Enabled = False
134             txtAddress.Enabled = False
136             txtStatus.Enabled = False
            
            Else
        
138             cmdAddCollector.Enabled = True
140             cmdEdit.Enabled = False
142             cmdReplaceCollector.Enabled = False
144             cmdDelete.Enabled = False
            
146             txtFirstName.Enabled = True
148             txtLastname.Enabled = True
150             txtMI.Enabled = True
152             txtAddress.Enabled = True
154             txtStatus.Enabled = True
            End If
                
156         rsCollData.MoveFirst
158         txtLastname.Text = rsCollData!lastname
160         txtFirstName.Text = rsCollData!firstname
162         txtMI.Text = rsCollData!MI
164         txtAddress.Text = rsCollData!Adress
166         dtEmployed.Value = rsCollData!DateEmployed
168         lblCollectorID.Caption = rsCollData!ID
170         dtEmployed.Enabled = False
        
172         If rsCollData3.State = 1 Then rsCollData3.Close
174         rsCollData3.Open "Select * from tblColl_Data where Code = " & _
                    cmbCollCode2.Text & " order by DateEmployed DESC"
        
176         If rsCollData3.RecordCount <> 0 Then
178             rsCollData3.MoveFirst

180             If rsCollData3.RecordCount = 1 Then
182                 txtStatus.Text = "Active"
184             ElseIf lblCollectorID.Caption = rsCollData3!ID Then
186                 txtStatus.Text = "Active"
                Else
188                 txtStatus.Text = "Replaced"
                End If
            End If

        Else
190         Call clear_data
        End If

        'Benjamin sumilhig
        'Mao ni ang mo check if ang collector nga iya e switch naka received na ba ug payment ana nga day
        'if nakareceived na gani dle na siya pwede e switch para dle maapektohan ang sub recport nga collection sa DCR
192     If cmdAddCollector.Caption = "Confirm Switch" Then
194         If rsPayment.State = 1 Then rsPayment.Close
196         rsPayment.Open "SELECT * FROM tblPayment WHERE CollectorCode = " & _
                    cmbCollCode2.Text & " AND Date = #" & Format$(DateValue(Now), _
                    "mm/dd/yy") & "# AND (Status = 'GOOD' OR Status = 'Fully Paid') "

198         If rsPayment.RecordCount > 0 Then
                'clear frame label
200             framePayment_made.Height = 1335
202             LabelCollectInfo.Caption = " "
204             label_orandname.Caption = " "

                'Get the collector name
206             If rsCollData.State = 1 Then rsCollData.Close
208             rsCollData.Open "SELECT * FROM tblColl_Data WHERE Code = '" & _
                        cmbCollCode2.Text & "' ORDER BY DateEmployed Desc"
210             LabelCollectInfo.Caption = "Code : " & cmbCollCode2.Text & " Name: " & _
                        rsCollData!lastname & " " & rsCollData!firstname & _
                        ". This collector is not available for switching area code."
212             lblMessage.Caption = _
                        "Because of the payment recieved, if you want to switch this collector please reverse all the payments below."

214             If rsPayment.State = 1 Then rsPayment.Close
216             rsPayment.Open _
                        "SELECT * FROM tblPayment WHERE Code <> 'Over' AND CollectorCode = '" _
                        & cmbCollCode2.Text & "' AND Date = #" & Format$(DateValue( _
                        Now), "mm/dd/yy") & _
                        "# AND (Status = 'GOOD' OR Status = 'Fully Paid') "

218             If rsPayment.RecordCount <> 0 Then

220                 Do While Not rsPayment.EOF
222                     framePayment_made.Height = framePayment_made.Height + 285
                        'get the complete name of the customer that will display in the frame of  payments made if the collector have a recieved a payment during the
                        'swicthing of the collector
224                     Call get_name(rsPayment!ORnumber, rsPayment!code, _
                                rsPayment!paymentsMade)
226                     rsPayment.MoveNext
                    Loop
                
                End If

                'disabled the combobox
228             cmdAddCollector.Enabled = False
230             cmbCollCode.Enabled = False
232             framePayment_made.Visible = True
            Else
234             cmdAddCollector.Enabled = True
236             cmbCollCode.Enabled = True
238             framePayment_made.Visible = False
            End If
        End If

        '<EhFooter>

        TxtLog "Exited cmbCollCode2_Click"

        Exit Sub

cmbCollCode2_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.cmbCollCode2_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmbCollCode_Click()

        '<EhHeader>
        On Error GoTo cmbCollCode_Click_Err

        TxtLog "Entered cmbCollCode_Click"

        '</EhHeader>

        'Benjamin Sumilhig 3-12-2015 3:2pm
100     If cmbCollCode.Text > 1 Then
102         If cmbCollCode.Text = cmbCollCode2.Text Then
                'MsgBox (cmbCollCode.ListIndex)s
104             MsgBox ("Cannot switch with the same area")
        
                Dim cmbCount As Integer
        
106             For cmbCount = 0 To cmbCollCode.ListCount - 1
108                 cmbCollCode.ListIndex = cmbCount

110                 If cmbCollCode.Text <> cmbCollCode2.Text Then

                        Exit For

                    End If

112             Next cmbCount

            End If
        End If
    
114     If rsCollData.State = 1 Then rsCollData.Close
116     rsCollData.Open "Select * from tblColl_Data where Code = " & cmbCollCode.Text _
                & " order by DateEmployed DESC"

118     If rsCollData.RecordCount <> 0 Then
    
120         If cmdAddCollector.Caption <> "Confirm Switch" Then
        
122             cmdAddCollector.Enabled = False
124             cmdEdit.Enabled = True
126             cmdReplaceCollector.Enabled = True
128             cmdDelete.Enabled = True
            
130             txtFirstName.Enabled = False
132             txtLastname.Enabled = False
134             txtMI.Enabled = False
136             txtAddress.Enabled = False
138             txtStatus.Enabled = False
            Else
        
140             cmdAddCollector.Enabled = True
142             cmdEdit.Enabled = False
144             cmdReplaceCollector.Enabled = False
146             cmdDelete.Enabled = False
            
148             txtFirstName.Enabled = True
150             txtLastname.Enabled = True
152             txtMI.Enabled = True
154             txtAddress.Enabled = True
156             txtStatus.Enabled = True
            End If
                
158         rsCollData.MoveFirst
160         txtLastname.Text = rsCollData!lastname
162         txtFirstName.Text = rsCollData!firstname
164         txtMI.Text = rsCollData!MI
166         txtAddress.Text = rsCollData!Adress
168         dtEmployed.Value = rsCollData!DateEmployed
170         lblCollectorID.Caption = rsCollData!ID
172         dtEmployed.Enabled = False
        
174         If rsCollData3.State = 1 Then rsCollData3.Close
176         rsCollData3.Open "Select * from tblColl_Data where Code = " & _
                    cmbCollCode.Text & " order by DateEmployed DESC"
        
178         If rsCollData3.RecordCount <> 0 Then
180             rsCollData3.MoveFirst

182             If rsCollData3.RecordCount = 1 Then
184                 txtStatus.Text = "Active"
186             ElseIf lblCollectorID.Caption = rsCollData3!ID Then
188                 txtStatus.Text = "Active"
                Else
190                 txtStatus.Text = "Replaced"
                End If
            End If

        Else
192         Call clear_data
        End If

        'Benjamin sumilhig 03-11-2015
        'Mao ni ang mo check if ang collector nga iya e switch naka received na ba ug payment ana nga day
        'if nakareceived na gani dle na siya pwede e switch para dle maapektohan ang sub recport nga collection sa DCR

194     If cmdAddCollector.Caption = "Confirm Switch" Then
196         If rsPayment.State = 1 Then rsPayment.Close
198         rsPayment.Open "SELECT * FROM tblPayment WHERE CollectorCode = " & _
                    cmbCollCode.Text & " AND Date = #" & Format$(DateValue(Now), _
                    "mm/dd/yy") & _
                    "# AND (Status = 'GOOD' OR Status = 'Fully Paid') AND Code <> -1"

200         If rsPayment.RecordCount > 0 Then
                'clear frame label
202             framePayment_made.Height = 1335
204             LabelCollectInfo.Caption = " "
206             label_orandname.Caption = " "

                'Get the collector name
208             If rsCollData.State = 1 Then rsCollData.Close
210             rsCollData.Open "SELECT * FROM tblColl_Data WHERE Code = '" & _
                        cmbCollCode.Text & "' ORDER BY DateEmployed Desc"
212             LabelCollectInfo.Caption = "Code : " & cmbCollCode.Text & " Name: " & _
                        rsCollData!lastname & " " & rsCollData!firstname & _
                        ". This collector is not available for switching area code."
214             lblMessage.Caption = _
                        "Because of the payment recieved, if you want to switch this collector please reverse all the payments below."

216             If rsPayment.State = 1 Then rsPayment.Close
218             rsPayment.Open _
                        "SELECT * FROM tblPayment WHERE Code <> -1 AND CollectorCode = " _
                        & cmbCollCode.Text & " AND Date = #" & Format$(DateValue(Now), _
                        "mm/dd/yy") & _
                        "# AND (Status = 'GOOD' OR Status = 'Fully Paid') "

220             If rsPayment.RecordCount <> 0 Then

222                 Do While Not rsPayment.EOF
224                     framePayment_made.Height = framePayment_made.Height + 285
                        'get the complete name of the customer that will display in the frame of  payments made if the collector have a recieved a payment during the
                        'swicthing of the collector
226                     Call get_name(rsPayment!ORnumber, rsPayment!code, _
                                rsPayment!paymentsMade)
228                     rsPayment.MoveNext
                    Loop
                
                End If

                'disabled the combobox
230             cmdAddCollector.Enabled = False
232             cmbCollCode2.Enabled = False
234             framePayment_made.Visible = True
            Else
236             cmdAddCollector.Enabled = True
238             cmbCollCode2.Enabled = True
240             framePayment_made.Visible = False
            End If
        End If

        '<EhFooter>

        TxtLog "Exited cmbCollCode_Click"

        Exit Sub

cmbCollCode_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.cmbCollCode_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdAddCollector_Click()

        '<EhHeader>
        On Error GoTo cmdAddCollector_Click_Err

        TxtLog "Entered cmdAddCollector_Click"

        '</EhHeader>
        
100     If cmdAddCollector.Caption = "Add Collector Code" Then
    
102         cmdEdit.Enabled = False
104         cmdReplaceCollector.Enabled = False
106         dtEmployed.Enabled = True
108         txtFirstName.Enabled = True
110         txtLastname.Enabled = True
112         txtMI.Enabled = True
114         txtAddress.Enabled = True
116         txtStatus.Enabled = False
        
118         txtFirstName.Text = ""
120         txtLastname.Text = ""
122         txtMI.Text = ""
124         txtAddress.Text = ""
126         txtStatus.Text = ""
        
            Dim codetest    As String

            Dim codetestval As Integer
    
128         If rsCollCode.State = 1 Then rsCollCode.Close
130         rsCollCode.Open "Select * from tblColl_Code order by ID desc"
        
132         If rsCollCode.RecordCount <> 0 Then
134             rsCollCode.MoveFirst
                'cmbCollCode.Text = rsCollCode!Code
136             codetestval = Val(rsCollCode!code) + 1
            Else
138             codetestval = "1"
            End If
        
140         codetest = codetestval
142         cmbCollCode.Clear
144         cmbCollCode.AddItem codetest
146         cmbCollCode.ListIndex = 0
        
            'disable the other buttons
    
148         Call clear_data
    
150         cmdAddCollector.Caption = "Confirm Add"
152     ElseIf cmdAddCollector.Caption = "Confirm Add" Then

154         If rsCollData.State = 1 Then rsCollData.Close
156         rsCollData.Open "Select * FROM tblColl_Data WHERE FirstName = '" & Trim$( _
                    txtFirstName.Text) & "' AND LastName = '" & Trim$(txtLastname.Text) _
                    & "'"

158         If rsCollData.RecordCount > 0 Then
160             MsgBox "This Collector is already exist. Please choose another collector."
            Else

162             If Trim$(txtFirstName.Text) = "" And Trim$(txtLastname.Text) = "" And _
                        Trim$(txtMI.Text) = "" And Trim$(txtAddress.Text) = "" Then
164                 MsgBox ("Cannot add collector code without collector")

166             ElseIf Trim$(txtFirstName.Text) = "" Or Trim$(txtLastname.Text) = "" Or _
                        Trim$(txtMI.Text) = "" Or Trim$(txtAddress.Text) = "" Then
168                 MsgBox ("All Fields are required")
170             ElseIf MsgBox("Continue adding collector code?", vbQuestion + vbYesNo, _
                        "Webplus Lending Corporation") = vbYes Then

172                 If rsCollCode.State = 1 Then rsCollCode.Close
174                 rsCollCode.Open "Select * from tblColl_Code"

                    'add coll code
                
176                 With rsCollCode
178                     .AddNew
180                     !code = cmbCollCode.Text
182                     .Update
                    End With
                
                    'add coll data
                
184                 If rsCollData.State = 1 Then rsCollData.Close
186                 rsCollData.Open "Select * from tblColl_Data"
                
188                 With rsCollData
190                     .AddNew
192                     !code = cmbCollCode.Text
194                     !firstname = Trim$(txtFirstName.Text)
196                     !lastname = Trim$(txtLastname.Text)
198                     !MI = Trim$(txtMI.Text)
200                     !Adress = Trim$(txtAddress.Text)
202                     !collection = 0
204                     !YTDCollection = 0
206                     !DateEmployed = dtEmployed.Value
208                     .Update
                    End With

                    'Benjamin Sumilhig 3-3-2015
                    'Add tblCollector
210                 If rsCollector.State = 1 Then rsCollector.Close
212                 rsCollector.Open "Select * from tblCollector"
                
214                 With rsCollector
216                     .AddNew
218                     !code = cmbCollCode.Text
220                     !lastname = Trim$(txtLastname.Text)
222                     !firstname = Trim$(txtFirstName.Text)
224                     !MiddleInitial = Trim$(txtMI.Text)
226                     !Address = Trim$(txtAddress.Text)
228                     !DateToday = dtEmployed.Value
230                     !DateEmployed = dtEmployed.Value
232                     !collection = 0
234                     !YTDCollection = 0
236                     !BranchID = " "
238                     .Update
                    End With

                    'Benjamin Sumilhig 3-3-2015
240                 MsgBox ("Collector Code and Collector Successfully added")
242                 Call cmdRefresh_Click
                End If
            End If

244     ElseIf cmdAddCollector.Caption = "Confirm Switch" Then

246         If MsgBox("Are you sure?", vbQuestion + vbYesNo, _
                    "Webplus Lending Corporation") = vbYes Then
                'swapping codez here
                'Switch Code Added 3/3/2015 1:19 AM By: Jun Tavs
                'Mao ni ang code nga mo perform para mag swiching ug collectors
                'Mag ag collectors sa ilang area code
248             Call clear_data

250             If rsCollData5.State = 1 Then rsCollData5.Close
                'Query to Database for Collectors to switch
252             rsCollData5.Open "Select * from tblColl_Data where Code = '" & _
                        cmbCollCode2.Text & "' order by DateEmployed Desc"
                'junrey tavera calling sun procedures tua sa ubos palihug ko ug ukay
254             Call SwitchInfoTblCollector(cmbCollCode.Text, rsCollData5!firstname, _
                        rsCollData5!lastname, rsCollData5!MI, rsCollData5!Adress, _
                        rsCollData5!DateEmployed, rsCollData5!collection, _
                        rsCollData5!YTDCollection)

256             If rsCollData4.State = 1 Then rsCollData4.Close
258             rsCollData4.Open "Select * from tblColl_Data where Code = '" & _
                        cmbCollCode.Text & "' order by DateEmployed Desc"
                '03-16-2015 12:20am
                'junrey tavera calling sub procedure mao ra gehapun sa imabaw toa sa ubos palihug ko ug ukay
260             Call SwitchInfoTblCollector(cmbCollCode2.Text, rsCollData4!firstname, _
                        rsCollData4!lastname, rsCollData4!MI, rsCollData4!Adress, _
                        rsCollData4!DateEmployed, rsCollData4!collection, _
                        rsCollData4!YTDCollection)
                'Junrey Tavera Para ma switch pud ang collector sa mga loaner
               
262             If rsLoan1.State = 1 Then rsLoan1.Close
264             rsLoan1.Open "Select * FROM tblLoan Where Collector ='" & _
                        rsCollData5!lastname & "' AND CollectorFname = '" & _
                        rsCollData5!firstname & "'"

266             If rsLoan.State = 1 Then rsLoan.Close
268             rsLoan.Open "Select * FROM tblLoan Where Collector ='" & _
                        rsCollData4!lastname & "' AND CollectorFname = '" & _
                        rsCollData4!firstname & "'"

270             If rsLoan.RecordCount <> 0 Then

272                 Do While Not rsLoan.EOF

274                     With rsLoan
276                         !Collector = rsCollData5!lastname
278                         !CollectorFname = rsCollData5!firstname
280                         .Update
282                         .MoveNext
                        End With

                    Loop

                End If

284             If rsLoan1.RecordCount <> 0 Then

286                 Do While Not rsLoan1.EOF

288                     With rsLoan1
290                         !Collector = rsCollData4!lastname
292                         !CollectorFname = rsCollData4!firstname
294                         .Update
296                         .MoveNext
                        End With
                      
                    Loop

                End If

                'Declaring variable to hold rsCollData4
                Dim cCode As String
              
                'Assigening rsCollData's Value to Variable
298             cCode = rsCollData4!code
              
                'rsCollData5 switch to rsCollData4
300             With rsCollData4
302                 !code = rsCollData5!code
304                 .Update
                End With

                'rsCollData4 switch to rsCollData5 through assigend variable's
306             With rsCollData5
308                 !code = cCode
310                 .Update
                End With

312             MsgBox ("Swap successful")
314             Call cmdRefresh_Click
            End If
        End If

        '<EhFooter>

        TxtLog "Exited cmdAddCollector_Click"

        Exit Sub

cmdAddCollector_Click_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_Collectors.cmdAddCollector_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdClose_Click()

        '<EhHeader>
        On Error GoTo cmdClose_Click_Err

        TxtLog "Entered cmdClose_Click"

        '</EhHeader>

100     Unload Me

        '<EhFooter>

        TxtLog "Exited cmdClose_Click"

        Exit Sub

cmdClose_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.cmdClose_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdDelete_Click()

        '<EhHeader>
        On Error GoTo cmdDelete_Click_Err

        TxtLog "Entered cmdDelete_Click"

        '</EhHeader>

100     If cmdDelete.Caption = "Delete" Then
102         cmbCollCode.Enabled = False
104         cmdAddCollector.Enabled = False
106         cmdEdit.Enabled = False
108         cmdReplaceCollector.Enabled = False
        
110         cmdDelete.Caption = "Confirm Delete"
        Else
            'ugh checking again? haha.
        
112         If rsCollData3.State = 1 Then rsCollData3.Close
114         rsCollData3.Open "Select * from tblColl_Data where Code = '" & _
                    rsCollData!code & "' order by DateEmployed DESC"
        
116         If rsCollData3.RecordCount = 0 Then
118             MsgBox ("There are no collecors present for the collector code")
120         ElseIf rsCollData3.RecordCount = 1 Then
122             MsgBox ( _
                        "There is only one collector for the chosen collector area, please edit the collector information instead.")
124         ElseIf MsgBox("Are you sure you want to delete this collector?", vbQuestion _
                    + vbYesNo, "Webplus Lending Corporation") = vbYes Then

126             If rsCollData.State = 1 Then rsCollData.Close
128             rsCollData.Open "Delete * from tblColl_Data where ID = " & Val( _
                        lblCollectorID.Caption)
                '03-16-2015
                'junrey tavera function para ebalic niya ang information sa old nga collector sa tblCollector
130             Call RestCollectorInformation(rsCollData3!code)

                '03-16-2015
                'junrey tavera pag update sa tblLoan sa assign nga collector
132             If rsLoan.State = 1 Then rsLoan.Close
134             rsLoan.Open _
                        "Select Collector, LoanID, CollectorFname From tblloan WHERE Collector = '" _
                        & rsCollData4!lastname & "' AND CollectorFname = '" & _
                        rsCollData4!firstname & "' ORDER BY LoanID DESC"

136             If rsLoan.RecordCount <> 0 Then

138                 Do While Not rsLoan.EOF
140                     rsLoan!Collector = rsCollData!lastname
142                     rsLoan!CollectorFname = rsCollData!firstname
144                     rsLoan.Update
146                     rsLoan.MoveNext
                    Loop

                End If

148             Call cmdRefresh_Click
            End If
        End If

        '<EhFooter>

        TxtLog "Exited cmdDelete_Click"

        Exit Sub

cmdDelete_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.cmdDelete_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdEdit_Click()

        '<EhHeader>
        On Error GoTo cmdEdit_Click_Err

        TxtLog "Entered cmdEdit_Click"

        '</EhHeader>

100     If cmdEdit.Caption = "Edit" Then
102         cmbCollCode.Enabled = False
    
104         cmdAddCollector.Enabled = False
106         cmdReplaceCollector.Enabled = False
108         cmdDelete.Enabled = False

110         txtFirstName.Enabled = True
112         txtLastname.Enabled = True
114         txtMI.Enabled = True
116         txtAddress.Enabled = True
118         txtStatus.Enabled = False
        
120         dtEmployed.Enabled = True
        
122         cmdEdit.Caption = "Confirm Edit"
        Else
    
124         If rsCollData.State = 1 Then rsCollData.Close
126         rsCollData.Open "Select * from tblColl_Data where Code = '" & _
                    cmbCollCode.Text & "' order by DateEmployed DESC"
    
            Dim DateEm As Date

128         If rsCollData.RecordCount > 1 Then
130             rsCollData.MoveFirst
132             rsCollData.MoveNext
134             DateEm = rsCollData!DateEmployed
                '138             name = rsCollData!firstname
            Else
                '140             x = True
            End If
    
136         If DateEm >= dtEmployed.Value Then
138             MsgBox ( _
                        "Cannot set date employed to be less than or equal the date employed of the past collector")
140         ElseIf cmbCollCode.Text = "" Then
142             MsgBox ("Please select a collector code first") ' this filter does useless in this part. hahah.
144         ElseIf Trim$(txtFirstName.Text) = "" Or Trim$(txtLastname.Text) = "" Or _
                    Trim$(txtMI.Text) = "" Or Trim$(txtAddress.Text) = "" Then
146             MsgBox ("There are missing fields in the collector data")
                'ElseIf Condition Then
        
148         ElseIf MsgBox("Are you want to edit collector information?", vbQuestion + _
                    vbYesNo, "Webplus Lending Corporation") = vbYes Then
    
150             If rsCollData.State = 1 Then rsCollData.Close
152             rsCollData.Open "Select * from tblColl_Data where ID = " & Val( _
                        lblCollectorID.Caption)
                        
                'Benjamin Sumilhig 3-14-2015
                'MagStart ang Coding sa Pag edit sa information sa collector
                'Mo edit lang cia ug usbon ang firstname and lastname sa collector  kadtong tulo ka table tua gemention sa
                'pinaka ubos ani palihug lang ko ug pangita
            
154             If rsCollData!lastname <> Trim$(txtLastname.Text) Or _
                        rsCollData!firstname <> Trim$(txtFirstName.Text) Then

                    '3/13/2015 11:00 PM
                    'Jun Ray Tavera and Benjamin Sumilhig, Update Collector firstname and lastname in payment table
156                 If rsPayment.State = 1 Then rsPayment.Close
158                 rsPayment.Open _
                            "Select Collector, ID, CollectorFname From tblPayment Where Collector='" _
                            & rsCollData!lastname & "' AND CollectorFname = '" & _
                            rsCollData!firstname & "' ORDER BY ID DESC"

160                 If rsPayment.RecordCount <> 0 Then

162                     Do While Not rsPayment.EOF
164                         rsPayment!Collector = Trim$(txtLastname.Text)
166                         rsPayment!CollectorFname = Trim$(txtFirstName.Text)
168                         rsPayment.Update
170                         rsPayment.MoveNext
                        Loop

                    End If
                   
                    '3/13/2015 11:00 PM
                    'Jun Ray Tavera and Benjamin Sumilhig, Update Collector firstname and lastname in customer table
172                 If rsCustomer.State = 1 Then rsCustomer.Close
174                 rsCustomer.Open _
                            "Select Collector, ID, CollectorFirstname From tblCustomer WHERE Collector = '" _
                            & rsCollData!lastname & "' AND CollectorFirstname = '" & _
                            rsCollData!firstname & "' ORDER BY ID DESC"

176                 If rsCustomer.RecordCount <> 0 Then

178                     Do While Not rsCustomer.EOF
180                         rsCustomer!Collector = Trim$(txtLastname.Text)
182                         rsCustomer!CollectorFirstname = Trim$(txtFirstName.Text)
184                         rsCustomer.Update
186                         rsCustomer.MoveNext
                        Loop

                    End If

                    '3/13/2015 11:00 PM
                    'Jun Ray Tavera and Benjamin Sumilhig, Update Collector firstname and lastname in loan table
188                 If rsLoan.State = 1 Then rsLoan.Close
190                 rsLoan.Open _
                            "Select Collector, LoanID, CollectorFname From tblloan WHERE Collector = '" _
                            & rsCollData!lastname & "' AND CollectorFname = '" & _
                            rsCollData!firstname & "' ORDER BY LoanID DESC"

192                 If rsLoan.RecordCount <> 0 Then

194                     Do While Not rsLoan.EOF
196                         rsLoan!Collector = Trim$(txtLastname.Text)
198                         rsLoan!CollectorFname = Trim$(txtFirstName.Text)
200                         rsLoan.Update
202                         rsLoan.MoveNext
                        Loop

                    End If
                End If

                '3/16/2015 12:07 PM
                'Jun Ray Tavera edit to tblCollector
204             If rsCollector.State = 1 Then rsCollector.Close
206             rsCollector.Open "Select * From tblCollector Where Code = '" & _
                        rsCollData!code & "'"

208             With rsCollector
210                 !firstname = Trim$(txtFirstName.Text)
212                 !lastname = Trim$(txtLastname.Text)
214                 !MiddleInitial = Trim$(txtMI.Text)
216                 !Address = Trim$(txtAddress.Text)
218                 !DateEmployed = dtEmployed.Value
220                 .Update
                End With

                'end of editing the information of the collector tblcustomer tblpayment tblLoan
                    
222             If rsCollData.RecordCount <> 0 Then

224                 With rsCollData
226                     !firstname = Trim$(txtFirstName.Text)
228                     !lastname = Trim$(txtLastname.Text)
230                     !MI = Trim$(txtMI.Text)
232                     !Adress = Trim$(txtAddress.Text)
234                     !DateEmployed = dtEmployed.Value
236                     .Update
                    End With
    
238                 MsgBox ("collector info successfully edited")
                Else
240                 MsgBox ("there is an error excuting the query")
                End If
            
242             Call cmdRefresh_Click
            End If
        End If

        '<EhFooter>

        TxtLog "Exited cmdEdit_Click"

        Exit Sub

cmdEdit_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.cmdEdit_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdRefresh_Click()

        '<EhHeader>
        On Error GoTo cmdRefresh_Click_Err

        TxtLog "Entered cmdRefresh_Click"

        '</EhHeader>

100     Call get_current_collectors
102     Call load_coll_code2
104     Call load_coll_code
106     Call set_dp
    
108     cmbCollCode.Enabled = True
    
110     cmbCollCode2.Clear
112     cmbCollCode2.Enabled = False
114     cmbCollCode2.Visible = False
    
116     cmdAddCollector.Enabled = True
118     cmdAddCollector.Caption = "Add Collector Code"
120     cmdEdit.Enabled = False
122     cmdEdit.Caption = "Edit"
124     cmdReplaceCollector.Enabled = False
126     cmdReplaceCollector.Caption = "Replace Collector"
128     cmdDelete.Enabled = False
130     cmdDelete.Caption = "Delete"
132     cmdSwitchArea.Enabled = True
    
134     txtFirstName.Enabled = False
136     txtLastname.Enabled = False
138     txtMI.Enabled = False
140     txtAddress.Enabled = False
142     txtStatus.Enabled = False
    
144     dtEmployed.Enabled = False
146     dtSwitch.Enabled = False
148     dtSwitch.Visible = False
    
150     txtFirstName.Text = ""
152     txtLastname.Text = ""
154     txtMI.Text = ""
156     txtAddress.Text = ""
158     txtStatus.Text = ""
        'benjamin Sumilhig 3-14-2015
        'hinig click sa refresh kinahanglan ehide niya ang frame nga payments made
160     framePayment_made.Visible = False
162     Call setgrid

        '<EhFooter>

        TxtLog "Exited cmdRefresh_Click"

        Exit Sub

cmdRefresh_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.cmdRefresh_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdReplaceCollector_Click()

        '<EhHeader>
        On Error GoTo cmdReplaceCollector_Click_Err

        TxtLog "Entered cmdReplaceCollector_Click"

        '</EhHeader>

        'disable and enable stuff
100     If cmdReplaceCollector.Caption = "Replace Collector" Then
102         cmbCollCode.Enabled = False
    
104         cmdAddCollector.Enabled = False
106         cmdEdit.Enabled = False
108         cmdDelete.Enabled = False
        
110         txtFirstName.Enabled = True
112         txtLastname.Enabled = True
114         txtMI.Enabled = True
116         txtAddress.Enabled = True
118         txtStatus.Enabled = False
    
120         txtFirstName.Text = ""
122         txtLastname.Text = ""
124         txtMI.Text = ""
126         txtAddress.Text = ""
128         txtStatus.Text = ""
        
130         dtEmployed.Enabled = True
        
132         cmdReplaceCollector.Caption = "Confirm Replace"
        Else
    
134         If rsCollData.State = 1 Then rsCollData.Close
136         rsCollData.Open "Select * from tblColl_Data where Code = '" & _
                    cmbCollCode.Text & "' order by DateEmployed DESC"
        
            Dim FCollector As String

            Dim LCollector As String

138         FCollector = rsCollData!firstname
140         LCollector = rsCollData!lastname
        
            Dim DateEm As Date

            Dim x      As Boolean
    
142         x = False
    
144         If rsCollData.RecordCount <> 0 Then
146             rsCollData.MoveFirst
148             DateEm = rsCollData!DateEmployed
            Else
150             x = True
            End If
    
152         If x Then
154             MsgBox ("Something is wrong")
156         ElseIf DateEm >= dtEmployed.Value Then
158             MsgBox ( _
                        "Cannot set date employed to be less than or equal the date employed of the past collector")
160         ElseIf cmbCollCode.Text = "" Then
162             MsgBox ("Please select a collector code first")
164         ElseIf Trim$(txtFirstName.Text) = "" Or Trim$(txtLastname.Text) = "" Or _
                    Trim$(txtMI.Text) = "" Or Trim$(txtAddress.Text) = "" Then
166             MsgBox ("There are missing fields in the collector data")
                'ElseIf Condition Then
        
168         ElseIf MsgBox("Are you sure you want to replace collector?", vbQuestion + _
                    vbYesNo, "Webplus Lending Corporation") = vbYes Then
    
170             If rsCollData.State = 1 Then rsCollData.Close
172             rsCollData.Open "Select * from tblColl_Data"

                '03-16-2015 12:10 am
                'Mao ni ang mo update sa mga loaner under sa collector nga gereplace
174             If rsCollector.State = 1 Then rsCollector.Close
176             rsCollector.Open "Select * from tblCollector Where Code= '" & _
                        cmbCollCode.Text & "'"
            
178             If rsLoan.State = 1 Then rsLoan.Close
180             rsLoan.Open _
                        "Select Collector, LoanID, CollectorFname From tblloan WHERE Collector = '" _
                        & LCollector & "' AND CollectorFname = '" & FCollector & _
                        "' ORDER BY LoanID DESC"

182             If rsLoan.RecordCount <> 0 Then

184                 Do While Not rsLoan.EOF
186                     rsLoan!Collector = Trim$(txtLastname.Text)
188                     rsLoan!CollectorFname = Trim$(txtFirstName.Text)
190                     rsLoan.Update
192                     rsLoan.MoveNext
                    Loop

                End If
    
194             With rsCollData
196                 .AddNew
198                 !code = cmbCollCode.Text
200                 !firstname = Trim$(txtFirstName.Text)
202                 !lastname = Trim$(txtLastname.Text)
204                 !MI = Trim$(txtMI.Text)
206                 !Adress = Trim$(txtAddress.Text)
208                 !DateEmployed = dtEmployed.Value
210                 .Update
                End With

                '03-16-2015 12:10 am
                'Mao ni ang mo update sa tblcollector para dele ma outdated ang tblCollector
212             With rsCollector
214                 !firstname = Trim$(txtFirstName.Text)
216                 !lastname = Trim$(txtLastname.Text)
218                 !MiddleInitial = Trim$(txtMI.Text)
220                 !Address = Trim$(txtAddress.Text)
222                 !DateEmployed = dtEmployed.Value
224                 .Update
                End With
    
226             MsgBox ("collector successfully replaced")
            
228             Call cmdRefresh_Click
            End If
        End If

        '<EhFooter>

        TxtLog "Exited cmdReplaceCollector_Click"

        Exit Sub

cmdReplaceCollector_Click_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_Collectors.cmdReplaceCollector_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub cmdSwitchArea_Click()

        '<EhHeader>
        On Error GoTo cmdSwitchArea_Click_Err

        TxtLog "Entered cmdSwitchArea_Click"

        '</EhHeader>

100     If MsgBox( _
                "You are about to switch collector codes, are you sure you want to continue?", _
                vbQuestion + vbYesNo, "Webplus Lending Corporation") = vbYes Then
            'disable other buttons and stuff
102         Call cmdRefresh_Click
104         Call load_coll_code2
106         cmdAddCollector.Caption = "Confirm Switch"
108         cmdSwitchArea.Enabled = False
110         cmbCollCode2.Visible = True
112         cmbCollCode2.Enabled = True
114         dtSwitch.Enabled = True
116         dtSwitch.Visible = False
            '--Benjamin Sumilhig--------
118         cmdEdit.Enabled = False
120         cmdReplaceCollector.Enabled = False
122         cmdDelete.Enabled = False
            '
            '--------------------------
        
        End If

        '<EhFooter>

        TxtLog "Exited cmdSwitchArea_Click"

        Exit Sub

cmdSwitchArea_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.cmdSwitchArea_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DGCollectors_Click()

        '<EhHeader>
        On Error GoTo DGCollectors_Click_Err

        TxtLog "Entered DGCollectors_Click"

        '</EhHeader>
        
        '--------Benjamin refresh
100     Call load_coll_code2
102     cmbCollCode2.Clear
104     cmbCollCode2.Enabled = False
106     cmbCollCode2.Visible = False
108     dtSwitch.Enabled = False
110     dtSwitch.Visible = False
112     cmdSwitchArea.Enabled = False
        '------------------
114     cmbCollCode.Enabled = True
    
116     cmdAddCollector.Enabled = True
118     cmdAddCollector.Caption = "Add Collector Code"
120     cmdEdit.Enabled = False
122     cmdEdit.Caption = "Edit"
124     cmdReplaceCollector.Enabled = False
126     cmdReplaceCollector.Caption = "Replace Collector"
128     cmdDelete.Enabled = False
130     cmdDelete.Caption = "Delete"
    
132     txtFirstName.Enabled = False
134     txtLastname.Enabled = False
136     txtMI.Enabled = False
138     txtAddress.Enabled = False
140     txtStatus.Enabled = False
    
142     dtEmployed.Enabled = False
    
144     txtLastname.Text = rsCollData2!lastname
146     txtFirstName.Text = rsCollData2!firstname
148     txtMI.Text = rsCollData2!MI
150     txtAddress.Text = rsCollData2!Adress
152     dtEmployed.Value = rsCollData2!DateEmployed
154     lblCollectorID.Caption = rsCollData2!ID
    
        'cmbCollCode.Text = rsCollData2!Code
        'refresh
    
156     If rsCollData3.State = 1 Then rsCollData3.Close
158     rsCollData3.Open "Select * from tblColl_Data where Code = '" & rsCollData2!code _
                & "' order by DateEmployed DESC"
    
160     If rsCollData3.RecordCount <> 0 Then
162         rsCollData3.MoveFirst

164         If lblCollectorID.Caption = rsCollData3!ID Then
166             txtStatus.Text = "Active"
            Else
168             txtStatus.Text = "Replaced"
            End If
        End If

        '<EhFooter>

        TxtLog "Exited DGCollectors_Click"

        Exit Sub

DGCollectors_Click_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.DGCollectors_Click", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DGCollectors_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo DGCollectors_KeyPress_Err

        TxtLog "Entered DGCollectors_KeyPress"

        '</EhHeader>

100     KeyAscii = 0

        '<EhFooter>

        TxtLog "Exited DGCollectors_KeyPress"

        Exit Sub

DGCollectors_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_Collectors.DGCollectors_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DGCollectors_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, _
                                   Y As Single)

        '<EhHeader>
        On Error GoTo DGCollectors_MouseDown_Err

        TxtLog "Entered DGCollectors_MouseDown"

        '</EhHeader>

100     If Button = vbRightButton Then
102         DGCollectors.AllowUpdate = False
        Else
104         DGCollectors.AllowUpdate = True
        End If

        '<EhFooter>

        TxtLog "Exited DGCollectors_MouseDown"

        Exit Sub

DGCollectors_MouseDown_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_Collectors.DGCollectors_MouseDown", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        TxtLog "Entered Form_Load"

        '</EhHeader>

        'Benjamin Sumilhig 3-14-2015
        'hide frame payment

100     dtSwitch.Visible = False
102     framePayment_made.Visible = False
104     Call connect
106     Call payment
108     Call Loan
110     Call CollCode
112     Call CollData
        'Benjamin
114     Call Collector
        'Jun Ray
116     Call Customer
118     Call Customer2
120     Call Loan1
    
122     Call cmdRefresh_Click

        '<EhFooter>

        TxtLog "Exited Form_Load"

        Exit Sub

Form_Load_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.Form_Load", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtAddress_KeyPress_Err

        TxtLog "Entered txtAddress_KeyPress"

        '</EhHeader>

        Dim strvalid As String

        'Benjamin sumilhig
        'Nag add me ug special charater Para maka type siya ug Ò ug —

100     strvalid = _
                "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZÒ—0987654321/:;~`!@#$%^&*()_+=-'\.,|[]{} "

102     If (InStr(1, strvalid, Chr$(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 13) = 0 Then
104         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtAddress_KeyPress"

        Exit Sub

txtAddress_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.txtAddress_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtFirstname_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtFirstname_KeyPress_Err

        TxtLog "Entered txtFirstname_KeyPress"

        '</EhHeader>

        Dim strvalid As String

        'Benjamin sumilhig
        'Nag add me ug special charater Para maka type siya ug Ò ug —

100     strvalid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZÒ—/-' "

102     If (InStr(1, strvalid, Chr$(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 13) = 0 Then
104         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtFirstname_KeyPress"

        Exit Sub

txtFirstname_KeyPress_Err:
        ErrReport Err.Description, _
                "LendingClient.frm_Collectors.txtFirstname_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtLastname_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtLastname_KeyPress_Err

        TxtLog "Entered txtLastname_KeyPress"

        '</EhHeader>

        Dim strvalid As String

        'Benjamin sumilhig
        'Nag add me ug special charater Para maka type siya ug Ò ug —

100     strvalid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZÒ—/-' "

102     If (InStr(1, strvalid, Chr$(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 13) = 0 Then
104         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtLastname_KeyPress"

        Exit Sub

txtLastname_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.txtLastname_KeyPress", _
                Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtMI_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo txtMI_KeyPress_Err

        TxtLog "Entered txtMI_KeyPress"

        '</EhHeader>

        Dim strvalid As String

        'Benjamin sumilhig
        'Nag add me ug special charater Para maka type siya ug Ò ug —

100     strvalid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZÒ— "

102     If (InStr(1, strvalid, Chr$(KeyAscii)) Or KeyAscii = 8 Or KeyAscii = 13) = 0 Then
104         KeyAscii = 0
        End If

        '<EhFooter>

        TxtLog "Exited txtMI_KeyPress"

        Exit Sub

txtMI_KeyPress_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.txtMI_KeyPress", Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub txtsearch_Change()

        '<EhHeader>
        On Error GoTo txtsearch_Change_Err

        TxtLog "Entered txtsearch_Change"

        '</EhHeader>

100     If txtSearch.Text <> "" Then
102         If rsCollData4.State = 1 Then rsCollData4.Close
104         rsCollData4.Open "Select * from tblColl_Data where Code like '" & _
                    txtSearch.Text & "%' or FirstName like '" & txtSearch.Text & _
                    "%' or LastName like '" & txtSearch.Text & "%' or MI like '" & _
                    txtSearch.Text & "%'or DateEmployed like '" & txtSearch.Text & _
                    "%' order by DateEmployed DESC"
                
106         Set DGCollectors.DataSource = rsCollData4
108         DGCollectors.Columns(0).Width = 0
110         DGCollectors.Columns(1).Width = 700
112         DGCollectors.Columns(4).Width = 400
114         DGCollectors.Columns(5).Width = 3200
116         DGCollectors.Columns(7).Width = 0
118         DGCollectors.Refresh
        Else
120         Call setgrid
        End If

        '<EhFooter>

        TxtLog "Exited txtsearch_Change"

        Exit Sub

txtsearch_Change_Err:
        ErrReport Err.Description, "LendingClient.frm_Collectors.txtsearch_Change", Erl

        Resume Next

        '</EhFooter>

End Sub

