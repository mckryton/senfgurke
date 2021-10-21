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

Private m_reporter As TReport
Private m_start_time As Long
Private m_end_time As Long
Private m_statistics As TStatistics

Private Sub Class_Initialize()
    m_start_time = 0
    m_end_time = 0
    Set m_statistics = New TStatistics
End Sub

Private Sub Class_Terminate()
    Set m_statistics = Nothing
End Sub

Public Sub set_report_format(report_format As String)
    Select Case report_format
        Case "verbose", "v"
            reporter.Report_Formatter = New TReportFormatterVerbose
        Case "progress", "p"
            reporter.Report_Formatter = New TReportFormatterProgress
    End Select
End Sub

Public Property Get reporter() As TReport
    If m_reporter Is Nothing Then
        Set m_reporter = New TReport
    End If
    Set reporter = m_reporter
End Property

Public Property Let reporter(ByVal new_reporter As TReport)
    Set reporter = new_reporter
End Property

Public Sub run_test(Optional filter_tag, Optional feature_filter, Optional report_format, Optional features_as_text As Collection, Optional feature_dir)

    Dim parsed_features As Collection

    If IsMissing(filter_tag) Then filter_tag = vbNullString
    If IsMissing(report_format) Then report_format = "verbose"
    set_report_format CStr(report_format)
    If features_as_text Is Nothing Then Set features_as_text = TFeatureLoader.load_features(feature_dir, feature_filter, Me)
    Set parsed_features = parse_loaded_features(features_as_text)
    If parsed_features.Count > 0 Then run_parsed_features parsed_features, CStr(filter_tag)
    reporter.finish_report
    report_statistics EndTime - StartTime
    'ensure that this setting is refreshed even without any re-compile:
    TConfig.StepImplementations = Nothing
End Sub

Public Sub run_parsed_features(parsed_features As Collection, Optional filter_tag)

    Dim parsed_feature As Variant
    Dim feature As TFeature
    Dim feature_runner As TFeatureRunner
    
    If IsMissing(filter_tag) Then filter_tag = vbNullString
    Set feature_runner = New TFeatureRunner
    feature_runner.setup Me
    StartTime = ExtraVBA.get_unix_timestamp(Now, Timer)
    For Each parsed_feature In parsed_features
        'cast Variant to TFeature
        Set feature = parsed_feature
        feature_runner.run_feature feature, filter_tag:=CStr(filter_tag)
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

Public Function get_duration_as_text(duration_in_ms As Long) As String
    
    Dim minutes As Long
    Dim seconds As Long
    Dim milliseconds As Long
    
    minutes = CLng(Int(duration_in_ms / 60000))
    seconds = CLng(Int((duration_in_ms - (minutes * 60000)) / 1000))
    milliseconds = duration_in_ms - (minutes * 60000) - (seconds * 1000)
    get_duration_as_text = minutes & "m " & seconds & "." & Format(milliseconds, "000") & "s"
End Function

Public Function parse_loaded_features(feature_list As Collection) As Collection
    
    Dim loaded_feature As Variant
    Dim parsed_features As Collection
    Dim parsed_feature As TFeature

    Set parsed_features = New Collection
    On Error GoTo parse_error
    For Each loaded_feature In feature_list
        Set parsed_feature = parse_feature(CStr(loaded_feature("feature_text")))
        If Not parsed_feature Is Nothing Then parsed_features.Add parsed_feature
    Next
    Set parse_loaded_features = parsed_features
    Exit Function
    
parse_error:
    If Err.Number >= 6000 And Err.Number < 7000 Then
        reporter.report REPORT_MSG_TYPE_PARSE_ERR, "Found syntax error while parsing feature '" _
                        & loaded_feature("origin") & "'", err_msg:=Err.description
    End If
    Resume Next
End Function

Public Sub report_statistics(run_duration As Long)

    Dim report_msg As String
    
    report_msg = statistics.get_statistics_summary() & vbLf _
                    & get_duration_as_text(run_duration)
    reporter.report REPORT_MSG_TYPE_STATS, report_msg
End Sub

Public Property Get statistics() As TStatistics
    Set statistics = m_statistics
End Property

