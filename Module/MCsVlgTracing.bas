Attribute VB_Name = "MCsVlgTracing"
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : MCsVlgTracing
'    Project    : Project1
'
'    Description: [type_description_here]
'
'    Modified   :
'--------------------------------------------------------------------------------
'</CSCC>
' MCsVlgTracing - PLEASE DO NOT DELETE THIS LINE
'                 AND DO NOT ALTER THE METHOD NAMES IN ANY WAY

'CSEH: Skip

Option Explicit

' Specify the position where the tracing instructions are inserted
Public Enum ProcPosition

    ProcEnter
    ProcExit
    ProcInside

End Enum

' Variable which contains a reference to AxTools Visual Logger
Private m_oVLogger As VisualLogger

' Parameters for Visual Logger main call
Private m_Color    As ItemColor

Private m_Bold     As Boolean

Private m_Indent   As Long

Public Const E_ERR_CS_TRACING_INIT = vbObjectError + 1263

Private Const S_ERR_CS_TRACING_INIT = "Could not initiate tracing"

' Member      : AxCsDumpParamValue
' Description : Returns a formatted information about each method parameter
' Parameters  : sParamName    - Parameter name
'               vParamValue   - Parameter value
Public Function AxCsDumpParamValue(sParamName As String, vParamValue As Variant) As String

        Dim sRet$
    
100     If IsObject(vParamValue) Then
    
102         sRet = "Object"
        
104     ElseIf IsArray(vParamValue) Then
    
106         sRet = "Array"
        
        Else
    
            On Error Resume Next

108         sRet = CStr(vParamValue)
        
110         If Err.Number <> 0 Then
112             sRet = "Not determined"
            Else

114             If Len(sRet) > 13 Then sRet = Left$(sRet, 10) & "..."
            End If
        
        End If
    
116     AxCsDumpParamValue = sParamName & ": [" & sRet & "]"

End Function

' Member      : AxCsInitiateTrace
' Description : Instantiate the Visual Logger reference
Public Sub AxCsInitiateTrace()

        Static bErrReported As Boolean

        On Error GoTo hErr

        On Error Resume Next

        ' Instantiate a new Visual Logger client
100     Set m_oVLogger = New VisualLogger

        On Error GoTo hErr

102     If m_oVLogger Is Nothing Then
104         If Not bErrReported Then
106             bErrReported = True
108             Err.Raise E_ERR_CS_TRACING_INIT + 1, "AxCsInitiateTrace", _
                        "Could not create Visual Logger object. " & _
                        "Please check the AxTools Visual Logger COM server validity."
            End If

            Exit Sub

        End If

        ' Instantiate a new Visual Logger client
110     Set m_oVLogger = New VisualLogger
    
        ' Register the new client to the server
112     m_oVLogger.Register "Project1"
114     m_oVLogger.IndentSize = 1
    
        ' Initialize fields
116     m_Indent = 0
118     m_Color = -1
120     m_Bold = False
    
        Exit Sub

hErr:
122     Err.Raise E_ERR_CS_TRACING_INIT, "AxCsInitiateTrace", S_ERR_CS_TRACING_INIT

End Sub

' Member      : AxCsTerminateTrace
' Description : Used for cleaning purposes
Public Sub AxCsTerminateTrace()
    
100     Set m_oVLogger = Nothing
    
End Sub

' Member      : AxCsTrace
' Description : Send information to the Visual Logger window regarding the current method being processed
' Parameters  : ProjectName   - Name of the project which contains the method
'               ComponentName - Name of the component which contains the method
'               MemberName    - Name of the method being processed
'               TracePosition - Indicates the position within method body, either at start, at exit
'                               or inside it, where the AxCsTraceWatch method is called
' Notes       : For inside-member calls of AxCsTrace method, you can use the ProjectName
'               parameter to send tracing information.
Public Sub AxCsTrace(ByVal ProjectName As String, _
                     Optional ByVal ComponentName As String = "", _
                     Optional ByVal MemberName As String = "", _
                     Optional ByVal TracePosition As ProcPosition = ProcInside)

        Dim sTemp As String
    
        ' Make sure that tracing is initiated
100     If m_oVLogger Is Nothing Then AxCsInitiateTrace

        ' The color, bold state and suffix can be customized below
102     If TracePosition = ProcEnter Then
104         m_Color = Default
106         m_Bold = False
108         sTemp = " - enter"
110     ElseIf TracePosition = ProcExit Then
112         m_Color = Default
114         m_Bold = False
116         sTemp = " - exit"
        Else
118         m_Color = Default
120         m_Bold = False
122         sTemp = ""
        End If
    
        ' Decrease the indent for the current level of tracing information
124     If TracePosition = ProcExit Then
126         m_Indent = m_Indent - 4
        End If

128     If Not (TracePosition = ProcInside) Then
130         sTemp = ProjectName & "." & ComponentName & "." & MemberName & sTemp
        Else
132         sTemp = ProjectName
        End If
    
        ' Add tracing information to the Visual Logger window
        ' We have to use On Error Resume Next in order to prevent error -2147418107
        ' from COM/DCOM ("It is illegal to call out while inside message filter.")
        ' when tracing timers or message hooks.
        On Error Resume Next

134     m_oVLogger.AddEntry sTemp, m_Color, m_Bold, m_Indent
    
        ' Increase the indent for the next level of tracing information
136     If TracePosition = ProcEnter Then
138         m_Indent = m_Indent + 4
        End If
    
End Sub

' Member      : VLogger
' Description : Get the m_oVLogger object
Public Property Get VLogger() As VisualLogger
    
100     If m_oVLogger Is Nothing Then AxCsInitiateTrace

102     Set VLogger = m_oVLogger

End Property

