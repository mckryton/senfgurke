VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_reporter As TReport
Private m_start_time As Long
Private m_end_time As Long
Private m_statistics As TStatistics
Private m_step_implementations As Collection
Private m_event_logger As TEventLogger
Private m_execution_hooks As TExecutionHooks

Private Sub Class_Initialize()
    m_start_time = 0
    m_end_time = 0
    Set m_event_logger = New TEventLogger
    Set m_execution_hooks = New TExecutionHooks
    Set m_statistics = New TStatistics
    Set m_statistics.SessionEventBroker = m_event_logger
    Set m_reporter = New TReport
    Set m_reporter.SessionEventBroker = m_event_logger
End Sub

Private Sub Class_Terminate()
    Set m_reporter = Nothing
    Set m_statistics = Nothing
    Set m_step_implementations = Nothing
    Set m_event_logger = Nothing
End Sub

Public Property Get EventLogger() As TEventLogger
    'handle events for internal purposes, eg. collecting statistics
    Set EventLogger = m_event_logger
End Property

Public Property Get ExecutionHooks() As TExecutionHooks
    'provides an interface for test code, eg. to get notified whenever an example excution was complete
    Set ExecutionHooks = m_execution_hooks
End Property

Public Sub log_event(event_name As String, Optional event_msg, Optional section_name, Optional section_type, _
                        Optional section_status, Optional feature_file, Optional log_time)
    Dim new_event As TEvent
    
    Set new_event = New TEvent
    new_event.EventName = event_name
    If Not IsMissing(event_msg) Then new_event.EventMessage = event_msg
    If Not IsMissing(section_name) Then new_event.SectionName = section_name
    If Not IsMissing(section_type) Then new_event.SectionType = section_type
    If Not IsMissing(section_status) Then new_event.SectionStatus = CStr(section_status)
    If Not IsMissing(feature_file) Then new_event.FeatureFileName = feature_file
    If Not IsMissing(log_time) Then new_event.LogTime = CSng(log_time)
    m_event_logger.log_event new_event
    'trigger after event excution hooks for selected events
    call_execution_hook new_event
End Sub

Private Sub call_execution_hook(log_event As TEvent)
    Select Case log_event.EventName
        Case EVENT_RUN_EXAMPLE_FINISHED
            m_execution_hooks.call_after_example_hook log_event
    End Select
End Sub

Public Property Get StepImplementations() As Collection
    Dim step_implementation_class As Variant
    
    If m_step_implementations Is Nothing Then
        Set m_step_implementations = New Collection
    End If
    Set StepImplementations = m_step_implementations
End Property

Public Property Let StepImplementations(step_classes As Collection)
    Set m_step_implementations = step_classes
End Property

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

Public Sub run_test(step_classes As Collection, Optional filter_tag, Optional feature_filter, _
                        Optional report_format, Optional features_as_text As Collection, Optional application_dir As String)
    Dim parsed_features As Collection
    Dim feature_dir As String
    Dim statistic_summary As String

    Me.StepImplementations = step_classes
    If IsMissing(filter_tag) Then filter_tag = vbNullString
    If IsMissing(report_format) Then report_format = "verbose"
    If Not IsMissing(application_dir) Then
        If Right(application_dir, 1) <> ExtraVBA.get_path_separator Then
            application_dir = application_dir & ExtraVBA.get_path_separator
        End If
        feature_dir = application_dir & "features"
    Else
        feature_dir = vbNullString
    End If
    set_report_format CStr(report_format)
    'load feature from files only if no feature were given as text by parameter
    ' and a feature directory was specified
    If features_as_text Is Nothing And Not feature_dir = vbNullString Then
        Set features_as_text = TFeatureLoader.load_features(feature_dir, feature_filter, Me)
    End If
    Set parsed_features = parse_loaded_features(features_as_text)
    If parsed_features.Count > 0 Then run_parsed_features parsed_features, CStr(filter_tag)
    statistic_summary = statistics.get_statistics_summary() & vbLf & get_duration_as_text(EndTime - StartTime)
    Me.log_event EVENT_RUN_SESSION_FINISHED, statistic_summary
End Sub

Public Function get_duration_as_text(duration_in_ms As Long) As String
    
    Dim minutes As Long
    Dim seconds As Long
    Dim milliseconds As Long
    
    minutes = CLng(Int(duration_in_ms / 60000))
    seconds = CLng(Int((duration_in_ms - (minutes * 60000)) / 1000))
    milliseconds = duration_in_ms - (minutes * 60000) - (seconds * 1000)
    get_duration_as_text = minutes & "m " & seconds & "." & Format(milliseconds, "000") & "s"
End Function

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

Public Function parse_loaded_features(feature_list As Collection) As Collection
    Dim loaded_feature As Variant
    Dim parsed_features As Collection
    Dim parsed_feature As TFeature

    Set parsed_features = New Collection
    On Error GoTo parse_error
    For Each loaded_feature In feature_list
        Set parsed_feature = TFeatureParser.parse_feature(CStr(loaded_feature("feature_text")))
        parsed_feature.FeatureFileName = loaded_feature("origin")
        If Not parsed_feature Is Nothing Then parsed_features.Add parsed_feature
    Next
    Set parse_loaded_features = parsed_features
    Exit Function
    
parse_error:
    If Err.Number >= 6000 And Err.Number < 7000 Then
        Me.log_event EVENT_PARSE_ERROR, Err.description, loaded_feature("origin"), SECTION_TYPE_FEATURE
    End If
    Resume Next
End Function

Public Property Get statistics() As TStatistics
    Set statistics = m_statistics
End Property

Public Sub enable_silent_mode()
    'if called any reporting activity is suppressed
    Me.reporter.enable_silent_mode
End Sub
