VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_logger As Logger
Dim m_reporter As TReport
Dim m_start_time As Long
Dim m_end_time As Long
Dim m_run_statistics As Collection

Public Property Get Log() As Logger
    If m_logger Is Nothing Then
        Set m_logger = New Logger
    End If
    Set Log = m_logger
End Property

Public Sub set_report_format(report_format As String)
    Select Case report_format
        Case "verbose", "v"
            Reporter.Report_Formatter = New TReportFormatterVerbose
        Case "progress", "p"
            Reporter.Report_Formatter = New TReportFormatterProgress
    End Select
End Sub

Public Property Get Reporter() As TReport
    If m_reporter Is Nothing Then
        Set m_reporter = New TReport
    End If
    Set Reporter = m_reporter
End Property

Public Property Let Reporter(ByVal new_reporter As TReport)
    Set Reporter = new_reporter
End Property

Public Sub run_test(Optional filter_tag, Optional report_format, Optional features_as_text As Collection, Optional silent)

    Dim parsed_features As Collection

    If IsMissing(filter_tag) Then filter_tag = vbNullString
    If IsMissing(report_format) Then report_format = "verbose"
    set_report_format CStr(report_format)
    If features_as_text Is Nothing Then Set features_as_text = TFeatureLoader.load_features
    If IsMissing(silent) Then silent = False
    Set parsed_features = TFeatureParser.parse_loaded_features(features_as_text)
    run_features parsed_features, CStr(filter_tag), silent
    Reporter.finish_report silent
    report_statistics RunStatistics, EndTime - StartTime, silent
    TConfig.StepImplementations = Nothing
End Sub

Public Sub run_features(parsed_features As Collection, Optional filter_tag, Optional silent)

    Dim parsed_feature As Variant
    Dim feature As TFeature
    
    If IsMissing(filter_tag) Then filter_tag = vbNullString
    If IsMissing(silent) Then silent = False
    StartTime = ExtraVBA.get_unix_timestamp(Now, Timer)
    For Each parsed_feature In parsed_features
        'cast Variant to TFeature
        Set feature = parsed_feature
        RunStatistics.Add TFeatureRunner.run_feature(feature, filter_tag:=CStr(filter_tag), silent:=silent)
    Next
    EndTime = ExtraVBA.get_unix_timestamp(Now, Timer)
End Sub

Public Property Get StartTime() As Long
    StartTime = m_start_time
End Property

Public Property Let StartTime(ByVal new_start_time As Long)
    m_start_time = new_start_time
End Property

Public Property Get EndTime() As Long
    EndTime = m_end_time
End Property

Public Property Let EndTime(ByVal new_end_time As Long)
    m_end_time = new_end_time
End Property

Private Sub Class_Initialize()
    m_start_time = 0
    m_end_time = 0
End Sub

Public Function get_duration_as_text(duration_in_ms As Long) As String
    
    Dim formatted_duration As String
    Dim minutes As Long
    Dim seconds As Long
    Dim milliseconds As Long
    
    minutes = CLng(Int(duration_in_ms / 60000))
    seconds = CLng(Int((duration_in_ms - (minutes * 60000)) / 1000))
    milliseconds = duration_in_ms - (minutes * 60000) - (seconds * 1000)
    get_duration_as_text = minutes & "m " & seconds & "." & Format(milliseconds, "000") & "s"
End Function

Public Property Get RunStatistics() As Collection
    If m_run_statistics Is Nothing Then
        Set m_run_statistics = New Collection
    End If
    Set RunStatistics = m_run_statistics
End Property

Public Property Let RunStatistics(ByVal new_run_statistics As Collection)
    Set m_run_statistics = new_run_statistics
End Property

Public Function get_step_statistics(run_statistics As Collection) As String
    
    Dim index As Integer
    Dim states As Variant
    Dim step_statistics As Collection
    Dim feature_log As Variant
    Dim example_log As Variant
    Dim step_log As Variant
    Dim step_results As String
    Dim counter As Long
    
    Set step_statistics = New Collection
    states = Array(STATUS_OK, STATUS_FAIL, STATUS_MISSING, STATUS_PENDING, STATUS_SKIPPED)
    For index = 0 To UBound(states)
        step_statistics.Add 0, states(index)
    Next
    step_statistics.Add 0, "sum"
    For Each feature_log In run_statistics
        For Each example_log In feature_log
            For Each step_log In example_log
                If IsArray(step_log) Then
                    counter = step_statistics(step_log(0)) + 1
                    step_statistics.Remove step_log(0)
                    step_statistics.Add counter, step_log(0)
                    counter = step_statistics("sum") + 1
                    step_statistics.Remove "sum"
                    step_statistics.Add counter, "sum"
                End If
            Next
        Next
    Next
    step_results = step_statistics("sum") & " steps ("
    For index = 0 To UBound(states)
        If step_statistics(states(index)) > 0 Then
            step_results = step_results & step_statistics(states(index)) & " " & states(index) & ", "
        End If
    Next
    step_results = Replace(step_results, STATUS_OK, "passed")
    step_results = Replace(step_results, STATUS_FAIL, "failed")
    step_results = Replace(step_results, STATUS_MISSING, "undefined")
    step_results = Replace(step_results, STATUS_PENDING, "pending")
    step_results = Replace(step_results, STATUS_SKIPPED, "skipped")
    If Right(step_results, 2) = ", " Then step_results = Left(step_results, Len(step_results) - 2)
    If Right(step_results, 1) = "(" Then
        step_results = Left(step_results, Len(step_results) - 1)
    Else
        step_results = step_results & ")"
    End If
    get_step_statistics = step_results
End Function

Public Sub report_statistics(run_statistics As Collection, run_duration As Long, Optional silent)

    Dim report_msg As String
    
    If IsMissing(silent) Then silent = False
    report_msg = get_step_statistics(run_statistics) & vbLf _
                    & get_duration_as_text(run_duration)
    Reporter.report REPORT_MSG_TYPE_STATS, report_msg, silent:=silent
End Sub
