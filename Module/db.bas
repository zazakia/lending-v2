Attribute VB_Name = "DBsetup"
Option Explicit

'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : database defenition and setup
'    Project    : Loan System
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>

Public conn           As New adodb.Connection

Public rsLogin        As New adodb.Recordset

Public rsUser         As New adodb.Recordset

Public rsCustomer     As New adodb.Recordset

Public rsCustomer1    As New adodb.Recordset

Public rsCustomer2    As New adodb.Recordset

Public rsCollCode     As New adodb.Recordset

Public rsCollData     As New adodb.Recordset

Public rsCollData2    As New adodb.Recordset

Public rsCollData3    As New adodb.Recordset

Public rsCollData4    As New adodb.Recordset

Public rsCollData5    As New adodb.Recordset

Public rsCustomer_new As New adodb.Recordset

Public rsCollector    As New adodb.Recordset

Public rsDelivery     As New adodb.Recordset

Public rsCharge       As New adodb.Recordset

Public rsServicefee   As New adodb.Recordset

Public rsTrail        As New adodb.Recordset

Public rsLogtime      As New adodb.Recordset

Public rsLoan         As New adodb.Recordset

Public rsLoan1        As New adodb.Recordset

Public rsPayment      As New adodb.Recordset

Public rsPayment1     As New adodb.Recordset

Public rsCashOnHand   As New adodb.Recordset

Public rsCashOnBank   As New adodb.Recordset

Public rsAdjustment   As New adodb.Recordset

Public rsChart        As New adodb.Recordset

Public rsDeposit      As New adodb.Recordset

Public rsExpense      As New adodb.Recordset

Public rsBreak        As New adodb.Recordset

Public rsBranch       As New adodb.Recordset

Public Sub Adjustment()

        '<EhHeader>
        On Error GoTo Adjustment_Err

        TxtLog "Entered Adjustment"

        '</EhHeader>

100     Set rsAdjustment = Nothing
102     Set rsAdjustment = New adodb.Recordset
    
104     With rsAdjustment
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblAdjustment"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Adjustment"

        Exit Sub

Adjustment_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Adjustment", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Branch()

        '<EhHeader>
        On Error GoTo Branch_Err

        TxtLog "Entered Branch"

        '</EhHeader>

100     Set rsBranch = Nothing
102     Set rsBranch = New adodb.Recordset

104     With rsBranch
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblBranch"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Branch"

        Exit Sub

Branch_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Branch", Erl

        Resume Next

        '</EhFooter>

End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       LendingClient
' Procedure  :       Break
' Description:       for the money break down function database
' Created by :       Project Administrator
' Machine    :       LIGHT
' Date-Time  :       11/10/2017-7:21:45 PM
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Function Break()

        '<EhHeader>
        On Error GoTo Break_Err

        TxtLog "Entered Break"

        '</EhHeader>

100     Set rsBreak = Nothing
102     Set rsBreak = New adodb.Recordset
104     rsBreak.Open "Select * from tblBreakdown", conn, 1, 3

        '<EhFooter>

        TxtLog "Exited Break"

        Exit Function

Break_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Break", Erl

        Resume Next

        '</EhFooter>

End Function

Public Sub cashonbank()

        '<EhHeader>
        On Error GoTo cashonbank_Err

        TxtLog "Entered cashonbank"

        '</EhHeader>

100     Set rsCashOnBank = Nothing
102     Set rsCashOnBank = New adodb.Recordset

104     With rsCashOnBank
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblCashOnBank"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited cashonbank"

        Exit Sub

cashonbank_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.cashonbank", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Cashonhand()

        '<EhHeader>
        On Error GoTo Cashonhand_Err

        TxtLog "Entered Cashonhand"

        '</EhHeader>

100     Set rsCashOnHand = Nothing
102     Set rsCashOnHand = New adodb.Recordset

104     With rsCashOnHand
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblCashOnHand "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Cashonhand"

        Exit Sub

Cashonhand_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Cashonhand", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Charge()

        '<EhHeader>
        On Error GoTo Charge_Err

        TxtLog "Entered Charge"

        '</EhHeader>

100     Set rsCharge = Nothing
102     Set rsCharge = New adodb.Recordset

104     With rsCharge
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblCharge"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Charge"

        Exit Sub

Charge_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Charge", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Chart()

        '<EhHeader>
        On Error GoTo Chart_Err

        TxtLog "Entered Chart"

        '</EhHeader>

100     Set rsChart = Nothing
102     Set rsChart = New adodb.Recordset

104     With rsChart
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblChartOfAccounts"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Chart"

        Exit Sub

Chart_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Chart", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub CollCode()

        '<EhHeader>
        On Error GoTo CollCode_Err

        TxtLog "Entered CollCode"

        '</EhHeader>

100     Set rsCollCode = Nothing
102     Set rsCollCode = New adodb.Recordset
 
104     With rsCollCode
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblColl_Code"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited CollCode"

        Exit Sub

CollCode_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.CollCode", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub CollData()

        '<EhHeader>
        On Error GoTo CollData_Err

        TxtLog "Entered CollData"

        '</EhHeader>

100     Set rsCollData = Nothing
102     Set rsCollData = New adodb.Recordset
 
104     With rsCollData
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblColl_Data"
114         .CursorLocation = adUseClient
116         .Open
        End With
        
118     Set rsCollData2 = Nothing
120     Set rsCollData2 = New adodb.Recordset
 
122     With rsCollData2
124         .CursorType = adOpenDynamic
126         .LockType = adLockOptimistic
128         .ActiveConnection = conn
130         .Source = "Select * from tblColl_Data"
132         .CursorLocation = adUseClient
134         .Open
        End With
        
136     Set rsCollData3 = Nothing
138     Set rsCollData3 = New adodb.Recordset
 
140     With rsCollData3
142         .CursorType = adOpenDynamic
144         .LockType = adLockOptimistic
146         .ActiveConnection = conn
148         .Source = "Select * from tblColl_Data"
150         .CursorLocation = adUseClient
152         .Open
        End With
        
154     Set rsCollData4 = Nothing
156     Set rsCollData4 = New adodb.Recordset
        
158     With rsCollData4
160         .CursorType = adOpenDynamic
162         .LockType = adLockOptimistic
164         .ActiveConnection = conn
166         .Source = "Select * from tblColl_Data"
168         .CursorLocation = adUseClient
170         .Open
        End With
        
172     Set rsCollData5 = Nothing
174     Set rsCollData5 = New adodb.Recordset
        
176     With rsCollData5
178         .CursorType = adOpenDynamic
180         .LockType = adLockOptimistic
182         .ActiveConnection = conn
184         .Source = "Select * from tblColl_Data"
186         .CursorLocation = adUseClient
188         .Open
        End With

        '<EhFooter>

        TxtLog "Exited CollData"

        Exit Sub

CollData_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.CollData", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Collector()

        '<EhHeader>
        On Error GoTo Collector_Err

        TxtLog "Entered Collector"

        '</EhHeader>

100     Set rsCollector = Nothing
102     Set rsCollector = New adodb.Recordset
 
104     With rsCollector
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblCollector"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Collector"

        Exit Sub

Collector_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Collector", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub connect()

        '<EhHeader>
        On Error GoTo connect_Err

        TxtLog "Entered connect"

        '</EhHeader>

100     Set conn = New adodb.Connection
    
102     conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & _
                App.Path & "\DB\JCashdb.mdb" & "; Jet OLEDB:Database Password=kim123;"
104     conn.Open

        '<EhFooter>

        TxtLog "Exited connect"

        Exit Sub

connect_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.connect", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Customer()

        '<EhHeader>
        On Error GoTo Customer_Err

        TxtLog "Entered Customer"

        '</EhHeader>

100     Set rsCustomer = Nothing
102     Set rsCustomer = New adodb.Recordset

104     With rsCustomer
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select TOP 50 * from tblCustomer "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Customer"

        Exit Sub

Customer_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Customer", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Customer2()

        '<EhHeader>
        On Error GoTo Customer2_Err

        TxtLog "Entered Customer2"

        '</EhHeader>

100     Set rsCustomer2 = Nothing
102     Set rsCustomer2 = New adodb.Recordset

104     With rsCustomer2
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select TOP 50 * from tblCustomer "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Customer2"

        Exit Sub

Customer2_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Customer2", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub delivery()

        '<EhHeader>
        On Error GoTo delivery_Err

        TxtLog "Entered delivery"

        '</EhHeader>

100     Set rsDelivery = Nothing
102     Set rsDelivery = New adodb.Recordset

104     With rsDelivery
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblDelivery"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited delivery"

        Exit Sub

delivery_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.delivery", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Deposit()

        '<EhHeader>
        On Error GoTo Deposit_Err

        TxtLog "Entered Deposit"

        '</EhHeader>

100     Set rsDeposit = Nothing
102     Set rsDeposit = New adodb.Recordset

104     With rsDeposit
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblDeposit"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Deposit"

        Exit Sub

Deposit_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Deposit", Erl

        Resume Next

        '</EhFooter>

End Sub

'CSEH: ErrResumeNext
Public Sub ErrReport(sErrDesc As String, _
                     Optional sLocation As String = "", _
                     Optional iLine As Long = 0)

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    ' This routine is provided to be used in conjunction with the ErrReport error handling scheme
    ' It uses a global CAppSettings object (so you must insert the CAppSettings prebuilt component
    ' class in the project) that is assumed to be  initialized outside this routine, preferrably
    ' the same CAppSettings object used to store and retrieve your application's settings to and
    ' respectively from the system registry.
    '
    ' How to use: insert this routine in a module within your project or in a global class within
    ' your project or a referred project.

    Dim iFF%

    Dim bLog           As Boolean, bMsg As Boolean

    Static bNewSession As Boolean
    
    ' See if logging/msgbox is required/wanted
    Dim oAppSettings

    If (oAppSettings Is Nothing) Then

        bLog = True
        bMsg = True

    Else

        bLog = CBool(oAppSettings.GetSetting("General", "Logging", "True"))
        bMsg = CBool(oAppSettings.GetSetting("General", "ReportErrors", "False"))

    End If

    If bLog Then

        ' Logging required/wanted
        
        iFF = FreeFile
        Open App.Path & "\LogError.txt" For Append As #iFF
        Open App.Path & "\Log.txt" For Append As #iFF
        
        If Not bNewSession Then

            bNewSession = True
            Print #iFF, Date & "  - " & Time & " --- " & _
                    "New session....................................................."

        End If
        
        Print #iFF, Date & "  - " & Time & " --- " & sErrDesc & " --- in " & sLocation _
                & " / " & Str$(iLine)
        
        Close #iFF

    End If
    
    If bMsg Then

        ' MsgBox required/wanted

        'TODO: Replace the "MyAppName" string below with your application's name
        MsgBox "Error: " & sErrDesc & vbCrLf & vbCrLf & _
                "The error happened in component '" & sLocation & "' at line " & Trim$( _
                iLine) & " and was logged (if configured so) to the 'Log.txt' file.", _
                vbOKOnly + vbCritical, "Lending System"

    End If

End Sub

Public Sub Expense()

        '<EhHeader>
        On Error GoTo Expense_Err

        TxtLog "Entered Expense"

        '</EhHeader>

100     Set rsExpense = Nothing
102     Set rsExpense = New adodb.Recordset

104     With rsExpense
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblExpense"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Expense"

        Exit Sub

Expense_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Expense", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Loan()

        '<EhHeader>
        On Error GoTo Loan_Err

        TxtLog "Entered Loan"

        '</EhHeader>

100     Set rsLoan = Nothing
102     Set rsLoan = New adodb.Recordset

104     With rsLoan
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select TOP 50 * from tblLoan"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Loan"

        Exit Sub

Loan_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Loan", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Loan1()

        '<EhHeader>
        On Error GoTo Loan1_Err

        TxtLog "Entered Loan1"

        '</EhHeader>

100     Set rsLoan1 = Nothing
102     Set rsLoan1 = New adodb.Recordset

104     With rsLoan1
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select TOP 50 * from tblLoan "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Loan1"

        Exit Sub

Loan1_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Loan1", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Login()

        '<EhHeader>
        On Error GoTo Login_Err

        TxtLog "Entered Login"

        '</EhHeader>

100     Set rsLogin = Nothing
102     Set rsLogin = New adodb.Recordset

104     With rsLogin
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblUser"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Login"

        Exit Sub

Login_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Login", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Logtime()

        '<EhHeader>
        On Error GoTo Logtime_Err

        TxtLog "Entered Logtime"

        '</EhHeader>

100     Set rsLogtime = Nothing
102     Set rsLogtime = New adodb.Recordset

104     With rsLogtime
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblLogtime"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Logtime"

        Exit Sub

Logtime_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Logtime", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub payment()

        '<EhHeader>
        On Error GoTo payment_Err

        TxtLog "Entered payment"

        '</EhHeader>

100     Set rsPayment = Nothing
        
102     Set rsPayment = New adodb.Recordset

104     With rsPayment
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select TOP 50 * from tblPayment "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited payment"

        Exit Sub

payment_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.payment", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub payment1()

        '<EhHeader>
        On Error GoTo payment1_Err

        TxtLog "Entered payment1"

        '</EhHeader>

100     Set rsPayment1 = Nothing
        
102     Set rsPayment1 = New adodb.Recordset

104     With rsPayment1
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select TOP 50 * from tblPayment "
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited payment1"

        Exit Sub

payment1_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.payment1", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Servicefee()

        '<EhHeader>
        On Error GoTo Servicefee_Err

        TxtLog "Entered Servicefee"

        '</EhHeader>

100     Set rsServicefee = Nothing
102     Set rsServicefee = New adodb.Recordset

104     With rsServicefee
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblServicefee"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Servicefee"

        Exit Sub

Servicefee_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Servicefee", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub Trail()

        '<EhHeader>
        On Error GoTo Trail_Err

        TxtLog "Entered Trail"

        '</EhHeader>

100     Set rsTrail = Nothing
102     Set rsTrail = New adodb.Recordset

104     With rsTrail
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select TOP 50 * from tblTrail order by Date"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited Trail"

        Exit Sub

Trail_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.Trail", Erl

        Resume Next

        '</EhFooter>

End Sub

'CSEH: ErrResumeNext
Public Sub TxtLog(sText As String, Optional bNoDateTime As Boolean = False)

    '<EhHeader>
    On Error Resume Next

    '</EhHeader>

    ' This routine is provided to be used in conjunction with the ErrReportAndTrace error handling scheme
    ' as well as for any other tasks that require logging.
    '
    ' How to use: insert this routine in a module within your project or in a global class within
    ' your project or a referred project.

    Dim iFF%, sTrailer$

    Static bNewSession As Boolean

    sTrailer = ""

    If Not bNoDateTime Then sTrailer = Date & " - " & Time & " --- "
    
    iFF = FreeFile
    Open App.Path & "\Log.txt" For Append As #iFF
    'Open App.Path & "\LogError.txt" For Append As #iFF
        
    If Not bNewSession Then

        bNewSession = True
        Print #iFF, sTrailer & _
                "New session....................................................."

    End If

    Print #iFF, sTrailer & sText

    Close #iFF

End Sub

Public Sub UnloadAllForms(Optional FormToIgnore As String = "")

        '<EhHeader>
        On Error GoTo UnloadAllForms_Err

        TxtLog "Entered UnloadAllForms"

        '</EhHeader>

        Dim f As Form

100     For Each f In Forms

102         If f.Name <> FormToIgnore Then
104             Unload f
106             Set f = Nothing
            End If

108     Next f

        '<EhFooter>

        TxtLog "Exited UnloadAllForms"

        Exit Sub

UnloadAllForms_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.UnloadAllForms", Erl

        Resume Next

        '</EhFooter>

End Sub

Public Sub User()

        '<EhHeader>
        On Error GoTo User_Err

        TxtLog "Entered User"

        '</EhHeader>

100     Set rsUser = Nothing
102     Set rsUser = New adodb.Recordset

104     With rsUser
106         .CursorType = adOpenDynamic
108         .LockType = adLockOptimistic
110         .ActiveConnection = conn
112         .Source = "Select * from tblUser"
114         .CursorLocation = adUseClient
116         .Open
        End With

        '<EhFooter>

        TxtLog "Exited User"

        Exit Sub

User_Err:
        ErrReport Err.Description, "LendingClient.DBsetup.User", Erl

        Resume Next

        '</EhFooter>

End Sub

