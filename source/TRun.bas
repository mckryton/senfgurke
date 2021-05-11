Attribute VB_Name = "TRun"
'This is the starting point for all test runs. TRun.test will allow you to optional add tags and
' report formats while TRun.wip will call test with the @wip (work in progress) tag using the
' the verbose report format by default.

Option Explicit

Dim m_session As TSession

Public Sub test(Optional filter_tag, Optional report_format)
    Set m_session = New TSession
    session.run_test filter_tag, report_format
    Set m_session = Nothing
    cleanup_context
End Sub

Public Sub wip()
    'wip = work in progress
    test "@wip"
End Sub

Public Sub progress(Optional filter_tag)
    test filter_tag, report_format:="progress"
End Sub

Public Property Get session() As TSession
    Set session = m_session
End Property

Private Sub cleanup_context()
    'this is a workaround for using global variables
    ' solution: use a context object as session property to share values between step functions
    Set TStepVars.loaded_features = Nothing
End Sub
