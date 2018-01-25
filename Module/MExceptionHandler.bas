Attribute VB_Name = "MExceptionHandler"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MExceptionHandler
'    Project    : Project1
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
'CSEH: Skip
'*****************************************************************************************
'* Module      : MExceptionHandler
'* Description : Unhandled Exception Handler filter.
'* Notes       : This module is used to provide your application a handler for the
'*               unhandled exceptions that may cause your program to crash.
'* Source      : AxTools Source+ 2000 - The Source+ Library
'* Usage       : Call "InstallExceptionHandler" on project initialization (in Form_Load,
'*               Main etc.) and "UninstallExceptionHandler" on project cleanup (in
'*               Form_Unload, etc.)
'*****************************************************************************************

Option Explicit

' Private class API function declarations
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (pDest As EXCEPTION_RECORD, _
                                       ByVal LPEXCEPTION_RECORD As Long, _
                                       ByVal lngBytes As Long)

Private Declare Function SetUnhandledExceptionFilter _
                Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long

Private Declare Sub RaiseException _
                Lib "kernel32" (ByVal dwExceptionCode As Long, _
                                ByVal dwExceptionFlags As Long, _
                                ByVal nNumberOfArguments As Long, _
                                lpArguments As Long)

' Public enums
Public Enum EExceptionType

    eExceptionType_AccessViolation = &HC0000005
    eExceptionType_DataTypeMisalignment = &H80000002
    eExceptionType_Breakpoint = &H80000003
    eExceptionType_SingleStep = &H80000004
    eExceptionType_ArrayBoundsExceeded = &HC000008C
    eExceptionType_FaultDenormalOperand = &HC000008D
    eExceptionType_FaultDivideByZero = &HC000008E
    eExceptionType_FaultInexactResult = &HC000008F
    eExceptionType_FaultInvalidOperation = &HC0000090
    eExceptionType_FaultOverflow = &HC0000091
    eExceptionType_FaultStackCheck = &HC0000092
    eExceptionType_FaultUnderflow = &HC0000093
    eExceptionType_IntegerDivisionByZero = &HC0000094
    eExceptionType_IntegerOverflow = &HC0000095
    eExceptionType_PriviledgedInstruction = &HC0000096
    eExceptionType_InPageError = &HC0000006
    eExceptionType_IllegalInstruction = &HC000001D
    eExceptionType_NoncontinuableException = &HC0000025
    eExceptionType_StackOverflow = &HC00000FD
    eExceptionType_InvalidDisposition = &HC0000026
    eExceptionType_GuardPageViolation = &H80000001
    eExceptionType_InvalidHandle = &HC0000008
    eExceptionType_ControlCExit = &HC000013A

End Enum

' Private enums
Private Enum EExceptionHandlerReturn

    EXCEPTION_CONTINUE_EXECUTION = -1
    EXCEPTION_CONTINUE_SEARCH = 0
    EXCEPTION_EXECUTE_HANDLER = 1

End Enum

' Private class constants
'Maximum number of parameters an Exception_Record can have
Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15

' Private class type definitions
'Structure that contains processor-specific register data
Private Type CONTEXT

    FltF0        As Double
    FltF1        As Double
    FltF2        As Double
    FltF3        As Double
    FltF4        As Double
    FltF5        As Double
    FltF6        As Double
    FltF7        As Double
    FltF8        As Double
    FltF9        As Double
    FltF10       As Double
    FltF11       As Double
    FltF12       As Double
    FltF13       As Double
    FltF14       As Double
    FltF15       As Double
    FltF16       As Double
    FltF17       As Double
    FltF18       As Double
    FltF19       As Double
    FltF20       As Double
    FltF21       As Double
    FltF22       As Double
    FltF23       As Double
    FltF24       As Double
    FltF25       As Double
    FltF26       As Double
    FltF27       As Double
    FltF28       As Double
    FltF29       As Double
    FltF30       As Double
    FltF31       As Double
    
    IntV0        As Double
    IntT0        As Double
    IntT1        As Double
    IntT2        As Double
    IntT3        As Double
    IntT4        As Double
    IntT5        As Double
    IntT6        As Double
    IntT7        As Double
    IntS0        As Double
    IntS1        As Double
    IntS2        As Double
    IntS3        As Double
    IntS4        As Double
    IntS5        As Double
    IntFp        As Double
    IntA0        As Double
    IntA1        As Double
    IntA2        As Double
    IntA3        As Double
    IntA4        As Double
    IntA5        As Double
    IntT8        As Double
    IntT9        As Double
    IntT10       As Double
    IntT11       As Double
    IntRa        As Double
    IntT12       As Double
    IntAt        As Double
    IntGp        As Double
    IntSp        As Double
    IntZero      As Double
    
    Fpcr         As Double
    SoftFpcr     As Double
    
    Fir          As Double
    Psr          As Long
    
    ContextFlags As Long
    Fill(4)      As Long

End Type

'Structure that describes an exception.
Private Type EXCEPTION_RECORD

    ExceptionCode                                        As Long
    ExceptionFlags                                       As Long
    pExceptionRecord                                     As Long  ' Pointer to an EXCEPTION_RECORD structure
    ExceptionAddress                                     As Long
    NumberParameters                                     As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS)   As Long

End Type

'Structure that contains exception information that can be used by a debugger.
Private Type EXCEPTION_DEBUG_INFO

    pExceptionRecord     As EXCEPTION_RECORD
    dwFirstChance        As Long

End Type

'The EXCEPTION_POINTERS structure contains an exception record with a
'machine-independent description of an exception and a context record
'with a machine-dependent description of the processor context at the
'time of the exception.
Private Type EXCEPTION_POINTERS

    pExceptionRecord     As EXCEPTION_RECORD
    ContextRecord        As CONTEXT

End Type

' Private variables for internal use
Private blnFilterInstalled As Boolean

'*****************************************************************************************
'* Function    : ExceptionHandler
'* Notes       : This function receives an exception code value and returns the text
'*               description of the exception.
'*               This function will be called when an unhandled exception occurs.It raises
'*               an error so that it can be trapped with an ON ERROR statement in the
'*               procedure that caused the exception.
'*****************************************************************************************
Public Function ExceptionHandler(ByRef ExceptionPtrs As EXCEPTION_POINTERS) As Long

        On Error Resume Next
    
        Dim Rec          As EXCEPTION_RECORD

        Dim strException As String
  
        'Get the current exception record.
100     Rec = ExceptionPtrs.pExceptionRecord
  
        'If Rec.pExceptionRecord is not zero, then it is a nested exception and
        'Rec.pExceptionRecord points to another EXCEPTION_RECORD structure.  Follow
        'the pointers back to the original exception.
102     Do Until Rec.pExceptionRecord = 0
            'A friendly declaration of CopyMemory.
104         CopyMemory Rec, Rec.pExceptionRecord, Len(Rec)
        Loop
  
        'Translate the exception code into a user-friendly string.
106     strException = GetExceptionDescription(Rec.ExceptionCode)
  
        'Raise an error to return control to the calling procedure.
        On Error GoTo 0

108     Err.Raise 10000, "ExceptionHandler", strException
End Function

'*****************************************************************************************
'* Function    : GetExceptionDescription
'* Notes       : Returns a string containing the description of the occurred exception.
'*****************************************************************************************
Public Function GetExceptionDescription(ExceptionType As EExceptionType) As String

        On Error Resume Next
    
        Dim strDescription As String
  
100     Select Case ExceptionType
        
            Case eExceptionType_AccessViolation
102             strDescription = "Access Violation"
        
104         Case eExceptionType_DataTypeMisalignment
106             strDescription = "Data Type Misalignment"
        
108         Case eExceptionType_Breakpoint
110             strDescription = "Breakpoint"
        
112         Case eExceptionType_SingleStep
114             strDescription = "Single Step"
        
116         Case eExceptionType_ArrayBoundsExceeded
118             strDescription = "Array Bounds Exceeded"
        
120         Case eExceptionType_FaultDenormalOperand
122             strDescription = "Float Denormal Operand"
        
124         Case eExceptionType_FaultDivideByZero
126             strDescription = "Divide By Zero"
        
128         Case eExceptionType_FaultInexactResult
130             strDescription = "Floating Point Inexact Result"
        
132         Case eExceptionType_FaultInvalidOperation
134             strDescription = "Invalid Operation"
        
136         Case eExceptionType_FaultOverflow
138             strDescription = "Float Overflow"
        
140         Case eExceptionType_FaultStackCheck
142             strDescription = "Float Stack Check"
        
144         Case eExceptionType_FaultUnderflow
146             strDescription = "Float Underflow"
        
148         Case eExceptionType_IntegerDivisionByZero
150             strDescription = "Integer Divide By Zero"
        
152         Case eExceptionType_IntegerOverflow
154             strDescription = "Integer Overflow"
        
156         Case eExceptionType_PriviledgedInstruction
158             strDescription = "Privileged Instruction"
        
160         Case eExceptionType_InPageError
162             strDescription = "In Page Error"
        
164         Case eExceptionType_IllegalInstruction
166             strDescription = "Illegal Instruction"
        
168         Case eExceptionType_NoncontinuableException
170             strDescription = "Non Continuable Exception"
        
172         Case eExceptionType_StackOverflow
174             strDescription = "Stack Overflow"
        
176         Case eExceptionType_InvalidDisposition
178             strDescription = "Invalid Disposition"
        
180         Case eExceptionType_GuardPageViolation
182             strDescription = "Guard Page Violation"
        
184         Case eExceptionType_InvalidHandle
186             strDescription = "Invalid Handle"
        
188         Case eExceptionType_ControlCExit
190             strDescription = "Control-C Exit"
        
192         Case Else
194             strDescription = "Unknown (&H" & Right$("00000000" & Hex$( _
                        ExceptionType), 8) & ")"
    
        End Select
    
196     GetExceptionDescription = strDescription
End Function

'*****************************************************************************************
'* Function    : GetExceptionName
'* Notes       : Returns a string containing the name of the occurred exception.
'*****************************************************************************************
Public Function GetExceptionName(ExceptionType As EExceptionType) As String

        On Error Resume Next
    
        Dim strDescription As String
  
100     Select Case ExceptionType
        
            Case eExceptionType_AccessViolation
102             strDescription = "EXCEPTION_ACCESS_VIOLATION"
        
104         Case eExceptionType_DataTypeMisalignment
106             strDescription = "EXCEPTION_DATATYPE_MISALIGNMENT"
        
108         Case eExceptionType_Breakpoint
110             strDescription = "EXCEPTION_BREAKPOINT"
        
112         Case eExceptionType_SingleStep
114             strDescription = "EXCEPTION_SINGLE_STEP"
        
116         Case eExceptionType_ArrayBoundsExceeded
118             strDescription = "EXCEPTION_ARRAY_BOUNDS_EXCEEDED"
        
120         Case eExceptionType_FaultDenormalOperand
122             strDescription = "EXCEPTION_FLT_DENORMAL_OPERAND"
        
124         Case eExceptionType_FaultDivideByZero
126             strDescription = "EXCEPTION_FLT_DIVIDE_BY_ZERO"
        
128         Case eExceptionType_FaultInexactResult
130             strDescription = "EXCEPTION_FLT_INEXACT_RESULT"
        
132         Case eExceptionType_FaultInvalidOperation
134             strDescription = "EXCEPTION_FLT_INVALID_OPERATION"
        
136         Case eExceptionType_FaultOverflow
138             strDescription = "EXCEPTION_FLT_OVERFLOW"
        
140         Case eExceptionType_FaultStackCheck
142             strDescription = "EXCEPTION_FLT_STACK_CHECK"
        
144         Case eExceptionType_FaultUnderflow
146             strDescription = "EXCEPTION_FLT_UNDERFLOW"
        
148         Case eExceptionType_IntegerDivisionByZero
150             strDescription = "EXCEPTION_INT_DIVIDE_BY_ZERO"
        
152         Case eExceptionType_IntegerOverflow
154             strDescription = "EXCEPTION_INT_OVERFLOW"
        
156         Case eExceptionType_PriviledgedInstruction
158             strDescription = "EXCEPTION_PRIVILEGED_INSTRUCTION"
        
160         Case eExceptionType_InPageError
162             strDescription = "EXCEPTION_IN_PAGE_ERROR"
        
164         Case eExceptionType_IllegalInstruction
166             strDescription = "EXCEPTION_ILLEGAL_INSTRUCTION"
        
168         Case eExceptionType_NoncontinuableException
170             strDescription = "EXCEPTION_NONCONTINUABLE_EXCEPTION"
        
172         Case eExceptionType_StackOverflow
174             strDescription = "EXCEPTION_STACK_OVERFLOW"
        
176         Case eExceptionType_InvalidDisposition
178             strDescription = "EXCEPTION_INVALID_DISPOSITION"
        
180         Case eExceptionType_GuardPageViolation
182             strDescription = "EXCEPTION_GUARD_PAGE_VIOLATION"
        
184         Case eExceptionType_InvalidHandle
186             strDescription = "EXCEPTION_INVALID_HANDLE"
        
188         Case eExceptionType_ControlCExit
190             strDescription = "EXCEPTION_CONTROL_C_EXIT"
        
192         Case Else
194             strDescription = "Unknown"
    
        End Select
    
196     GetExceptionName = strDescription
End Function

'*****************************************************************************************
'* Sub         : InstallExceptionHandler
'* Notes       : Installs the custom exception filter.
'*****************************************************************************************
Public Sub InstallExceptionHandler()

        On Error Resume Next
    
100     If Not blnFilterInstalled Then
102         Call SetUnhandledExceptionFilter(AddressOf ExceptionHandler)
104         blnFilterInstalled = True
        End If

End Sub

'*****************************************************************************************
'* Sub         : RaiseAnException
'* Notes       : Raises an exception of the specified type.
'*****************************************************************************************
Public Sub RaiseAnException(ExceptionType As EExceptionType)
    
        'The following line raises an exception
100     RaiseException ExceptionType, 0, 0, 0
End Sub

'*****************************************************************************************
'* Sub         : UninstallExceptionHandler
'* Notes       : Uninstalls the custom exception filter and restore the default one.
'*****************************************************************************************
Public Sub UninstallExceptionHandler()

        On Error Resume Next
    
100     If blnFilterInstalled Then
102         Call SetUnhandledExceptionFilter(0&)
104         blnFilterInstalled = False
        End If

End Sub

