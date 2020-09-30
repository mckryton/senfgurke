VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const cLVL_DEBUG = 100
Const cLVL_INFO = 90
Const cLVL_WARNING = 50
Const cLVL_ERROR = 30
Const cLVL_CRITICAL = 1

Const LINE_BREAK = vbCr & vbLf

'current log level - decreasing log level means decreasing amount of messages
Dim mLogLevel As Integer

Public Sub Log(pLogMsg As String, Optional pMsgLogLevel)

    Dim msg_log_level As Integer
    Dim log_msg As String
    
    If IsMissing(pMsgLogLevel) Then
        msg_log_level = LVL_INFO
    Else
        msg_log_level = pMsgLogLevel
    End If
   
    If msg_log_level <= Me.LogLevel Then
        log_msg = Time
        Select Case msg_log_level
            Case Me.LVL_DEBUG
                log_msg = log_msg & " debug:"
            Case Me.LVL_INFO
                log_msg = log_msg & " info:"
            Case Me.LVL_WARNING
                log_msg = log_msg & " warning:"
            Case Me.LVL_ERROR
                log_msg = log_msg & " error:"
            Case Me.LVL_CRITICAL
                log_msg = log_msg & " critical:"
            Case Else
                log_msg = log_msg & " custom(" & msg_log_level & "):"
        End Select
        log_msg = log_msg & " " & pLogMsg
        Debug.Print log_msg
    End If
End Sub

Public Sub log_function_error(pFunctionName As String, Optional pLogMsg As Variant)

    Dim log_msg As String
    
    log_msg = Time & " error:" & LINE_BREAK & _
                vbTab & "source:" & vbTab & Err.Source & LINE_BREAK & _
                vbTab & "caller:" & vbTab & pFunctionName & LINE_BREAK & _
                vbTab & "desc:" & vbTab & Err.description & LINE_BREAK
    If Not IsMissing(pLogMsg) Then
        log_msg = log_msg & vbTab & "custom:" & vbTab & pLogMsg
    End If
    Debug.Print log_msg
End Sub

Public Sub debug_log(pLogMsg As String)
    Me.Log pLogMsg, LVL_DEBUG
End Sub

Public Sub info_log(pLogMsg As String)
    Me.Log pLogMsg, LVL_INFO
End Sub

Public Sub warning_log(pLogMsg As String)
    Me.Log pLogMsg, LVL_WARNING
End Sub

Public Sub error_log(pLogMsg As String)
    Me.Log pLogMsg, LVL_ERROR
End Sub

Public Sub critical_log(pLogMsg As String)
    Me.Log pLogMsg, LVL_CRITICAL
End Sub

Public Property Get LogLevel() As Integer
    LogLevel = mLogLevel
End Property

Public Property Let LogLevel(ByVal pLogLevel As Integer)
    mLogLevel = pLogLevel
End Property

Private Sub Class_Initialize()
    Me.LogLevel = LVL_INFO
End Sub

'because classes don't support public constants we need to have properties to make them accessible
Public Property Get LVL_DEBUG() As Integer
    LVL_DEBUG = cLVL_DEBUG
End Property

'because classes don't support public constants we need to have properties to make them accessible
Public Property Get LVL_INFO() As Integer
    LVL_INFO = cLVL_INFO
End Property

'because classes don't support public constants we need to have properties to make them accessible
Public Property Get LVL_WARNING() As Integer
    LVL_WARNING = cLVL_WARNING
End Property

'because classes don't support public constants we need to have properties to make them accessible
Public Property Get LVL_ERROR() As Integer
    LVL_ERROR = cLVL_ERROR
End Property

'because classes don't support public constants we need to have properties to make them accessible
Public Property Get LVL_CRITICAL() As Integer
    LVL_CRITICAL = cLVL_CRITICAL
End Property

