Attribute VB_Name = "TRun"
Option Explicit

Dim m_session As TSession

Public Sub test(Optional filter_tag, Optional report_format)
    Session.run_test filter_tag, report_format
End Sub

Public Sub wip()
    'wip = work in progress
    Session.run_test "@wip"
    Session = Nothing
End Sub

Public Property Get Session() As TSession
    If m_session Is Nothing Then
        Set m_session = New TSession
    End If
    Set Session = m_session
End Property

Public Property Let Session(ByVal new_session As TSession)
    Set m_session = new_session
End Property
