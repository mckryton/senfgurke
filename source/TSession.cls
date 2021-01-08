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

Public Sub run_test(Optional filter_tag, Optional report_format)

    Dim features_as_text As Collection
    Dim parsed_features As Collection

    If IsMissing(filter_tag) Then filter_tag = vbNullString
    If IsMissing(report_format) Then report_format = "verbose"
    set_report_format CStr(report_format)
    Set features_as_text = TFeatureLoader.load_features
    Set parsed_features = TFeatureParser.parse_loaded_features(features_as_text)
    run_features parsed_features, CStr(filter_tag)
    Reporter.report_code_templates_for_missing_steps
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
        TFeatureRunner.run_feature feature, filter_tag:=CStr(filter_tag), silent:=silent
    Next
    EndTime = ExtraVBA.get_unix_timestamp(Now, Timer)
    If Not silent Then Reporter.report REPORT_MSG_TYPE_DURATION, CStr(get_duration_in_ms(StartTime, EndTime))
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

Public Function get_duration_in_ms(start_time As Long, end_time As Long) As String
    
    Dim formatted_duration As String
    Dim duration_in_ms As Long
    Dim minutes As Long
    Dim seconds As Long
    Dim milliseconds As Long
    
    duration_in_ms = end_time - start_time
    minutes = CLng(Int(duration_in_ms / 60000))
    seconds = CLng(Int((duration_in_ms - (minutes * 60000)) / 1000))
    milliseconds = duration_in_ms - (minutes * 60000) - (seconds * 1000)
    get_duration_in_ms = minutes & "m " & seconds & "." & Format(milliseconds, "000") & "s"
End Function
