VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private report_out As TReportOut
Private m_report_formatter As Variant
Private m_err_buffer As String
Private m_code_template_buffer As Collection

Public WithEvents SessionEventBroker As TEventLogger
Attribute SessionEventBroker.VB_VarHelpID = -1

Private Sub Class_Initialize()
    'send output to the immediate window by default
    Set report_out = New TReportOutImmediateWindow
End Sub

Private Sub Class_Terminate()
    Set report_out = Nothing
    Set m_code_template_buffer = Nothing
End Sub

Public Sub add_err_msg_to_buffer(err_msg As String)
    If Not m_err_buffer = vbNullString Then m_err_buffer = m_err_buffer & vbLf
    m_err_buffer = m_err_buffer & err_msg
End Sub

Public Sub add_code_template_to_buffer(code_template As String)
    Dim function_name As String
    
    If m_code_template_buffer Is Nothing Then Set m_code_template_buffer = New Collection
    function_name = Trim(Split(code_template, vbLf)(0))
    'ignore function parameters
    function_name = Left(function_name, InStr(function_name, "(") - 1)
    'detect template duplicates
    If Not ExtraVBA.collection_has_key(function_name, m_code_template_buffer) Then
        m_code_template_buffer.Add code_template, function_name
    End If
End Sub

Public Sub print_err_buffer()
    report_out.PrintOut vbLf & m_err_buffer
End Sub

Private Sub report(message_type As String, message_content As String, Optional Status, Optional err_msg, Optional feature_file_name)

    Dim msg_package As Collection
    
    If Not Report_Formatter Is Nothing Then
        Set msg_package = New Collection
        If IsMissing(Status) Then
            msg_package.Add vbNullString, REPORT_MSG_STATUS
        Else
            msg_package.Add Status, REPORT_MSG_STATUS
        End If
        If IsMissing(err_msg) Then
            msg_package.Add vbNullString, REPORT_MSG_ERR
        Else
            msg_package.Add err_msg, REPORT_MSG_ERR
        End If
        If IsMissing(feature_file_name) Then
            msg_package.Add vbNullString, REPORT_MSG_FEATURE_ORIGIN
        Else
            msg_package.Add feature_file_name, REPORT_MSG_FEATURE_ORIGIN
        End If
        Set msg_package = build_msg_package(message_type, message_content, Status, err_msg, feature_file_name)
        Report_Formatter.process_msg msg_package
        Set msg_package = Nothing
    End If
End Sub

Public Function build_msg_package(message_type As String, message_content As String, Optional msg_status, Optional err_msg, Optional feature_file_name) As Collection

    Dim msg_package As Collection
    
    If IsMissing(msg_status) Then msg_status = vbNullString
    If IsMissing(err_msg) Then err_msg = vbNullString
    Set msg_package = New Collection
    With msg_package
        .Add message_type, REPORT_MSG_TYPE
        .Add message_content, REPORT_MSG_CONTENT
        .Add msg_status, REPORT_MSG_STATUS
        .Add err_msg, REPORT_MSG_ERR
        .Add feature_file_name, REPORT_MSG_FEATURE_ORIGIN
    End With
    Set build_msg_package = msg_package
End Function

Public Property Get Report_Formatter() As Variant
    If Not IsObject(m_report_formatter) Then
        Set m_report_formatter = New TReportFormatterVerbose
        m_report_formatter.reporter = Me
    End If
    Set Report_Formatter = m_report_formatter
End Property

Public Property Let Report_Formatter(ByVal new_report_formatter As Variant)
    Set m_report_formatter = new_report_formatter
    m_report_formatter.reporter = Me
End Property

Public Sub print_code_templates_for_missing_steps()

    Dim code_templates As String
    
    code_templates = get_code_templates_for_missing_steps
    If code_templates <> vbNullString Then report_out.PrintOut code_templates
End Sub

Public Function get_code_templates_for_missing_steps() As String

    Dim report_output As String
    Dim template As Variant

    If m_code_template_buffer Is Nothing Then
        report_output = vbNullString
    Else
        report_output = "You can implement step definitions for undefined steps with these snippets:"
        For Each template In m_code_template_buffer
            report_output = report_output & vbLf & vbLf & CStr(template)
        Next
        report_output = report_output & vbLf & vbLf & "If the function is already in place, the cause for this message could be:" & vbLf _
                            & "* the step implementation class is not registered in the TStepRegister module" & vbLf _
                            & "* the code inside the step function tries to access a non-existing method or property"
                            
    End If
    get_code_templates_for_missing_steps = report_output
End Function

Public Property Get ConsoleOutput() As String
    'this is for testing and debugging only, because you can't read the content of the immediate window
    ConsoleOutput = report_out.ConsoleOutput
End Property

Public Property Get CodeTemplateBuffer() As Collection
    Set CodeTemplateBuffer = m_code_template_buffer
End Property

Public Sub print_report(report_output As String, Optional consecutive As Boolean)
    If IsMissing(consecutive) Then consecutive = False
    report_out.PrintOut report_output, consecutive
End Sub

Public Sub finish_report()
    print_err_buffer
    print_code_templates_for_missing_steps
End Sub

Public Sub enable_silent_mode()
    Set report_out = Nothing
    Set report_out = New TReportOutMockConsole
End Sub

Private Sub SessionEventBroker_EventLogged(new_event As TEvent)
    Select Case new_event.EventName
        Case EVENT_RUN_FEATURE_STARTED
            report REPORT_MSG_TYPE_FEATURE_NAME, new_event.SectionName
            report REPORT_MSG_TYPE_DESC, new_event.EventMessage
        Case EVENT_RUN_RULE_STARTED
            report REPORT_MSG_TYPE_RULE, new_event.SectionName
        Case EVENT_RUN_EXAMPLE_STARTED
            report REPORT_MSG_TYPE_EXAMPLE_TITLE, new_event.SectionName
        Case EVENT_RUN_STEP_FINISHED
            report REPORT_MSG_TYPE_STEP, new_event.SectionName, new_event.SectionStatus, _
                    new_event.EventMessage, new_event.FeatureFileName
        Case EVENT_RUN_OUTLINE_STEP_DECLARED
            report REPORT_MSG_TYPE_OUTLINE_STEP, new_event.SectionName, new_event.SectionStatus, _
                    new_event.EventMessage, new_event.FeatureFileName
        Case EVENT_PARSE_ERROR
            report REPORT_MSG_TYPE_PARSE_ERR, "Found syntax error while parsing feature '" _
                        & new_event.SectionName & "'", err_msg:=new_event.EventMessage
        Case EVENT_RUN_SESSION_FINISHED
            report REPORT_MSG_TYPE_STATS, new_event.EventMessage
            finish_report
        Case EVENT_STEP_IS_MISSING
            report REPORT_MSG_TYPE_CODE_TEMPLATE, new_event.EventMessage
        Case EVENT_RUN_OUTLINE_EXAMPLE_STARTED
            If TypeName(Me.Report_Formatter) = "TReportFormatterVerbose" Then
                Me.Report_Formatter.OutlineWasStarted = True
            End If
            report REPORT_MSG_TYPE_EXAMPLE_TITLE, new_event.SectionName
        Case EVENT_RUN_OUTLINE_EXAMPLE_FINISHED
            If TypeName(Me.Report_Formatter) = "TReportFormatterVerbose" Then
                Me.Report_Formatter.OutlineWasStarted = False
            End If
        Case EVENT_RUN_OUTLINE_TABLE_STARTED
            report REPORT_MSG_TYPE_OUTLINE_TABLEHEADER, new_event.SectionName & vbLf & new_event.EventMessage
        Case EVENT_RUN_OUTLINE_ROW_STARTED
            report REPORT_MSG_TYPE_OUTLINE_ROW, new_event.EventMessage
    End Select
End Sub
