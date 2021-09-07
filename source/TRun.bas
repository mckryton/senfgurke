Attribute VB_Name = "TRun"
'This is the starting point for all test runs. TRun.test will allow you to optional add tags and
' report formats while TRun.wip will call test with the @wip (work in progress) tag using the
' the verbose report format by default.

Option Explicit

Public Sub test(Optional filter_tag, Optional feature_filter, Optional report_format)

    Dim session As TSession
    
    Set session = New TSession
    session.run_test filter_tag, feature_filter, report_format
    Set session = Nothing
End Sub

Public Sub wip()
    'wip = work in progress
    test "@wip", report_format:="verbose"
End Sub

Public Sub progress(Optional filter_tag)
    test filter_tag, report_format:="progress"
End Sub
