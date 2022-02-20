VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_message As String
Private m_msg_type As String
Private m_log_time As Single
Private m_status As String

Private Sub Class_Initialize()
    m_status = STATUS_UNKNOWN
    m_log_time = ExtraVBA.get_unix_timestamp_now
End Sub

Public Property Get message() As String
    message = m_message
End Property

Public Property Let message(ByVal new_message As String)
    m_message = new_message
End Property

Public Property Get msg_type() As String
    msg_type = m_msg_type
End Property

Public Property Let msg_type(ByVal new_msg_type As String)
    m_msg_type = new_msg_type
End Property

Public Property Get log_time() As String
    log_time = m_log_time
End Property

Public Property Let log_time(ByVal new_log_time As String)
    m_log_time = new_log_time
End Property

Public Property Get status() As String
    status = m_status
End Property

Public Property Let status(ByVal new_status As String)
    m_status = new_status
End Property


