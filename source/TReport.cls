VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_report_formatter As Variant
Dim m_console_output As String
Dim m_err_buffer As String
Dim m_code_template_buffer As Collection

Public Sub add_err_msg_to_buffer(err_msg As String)
    If Not m_err_buffer = vbNullString Then m_err_buffer = m_err_buffer & vbLf
    m_err_buffer = m_err_buffer & err_msg
End Sub

Public Sub add_code_template_to_buffer(code_template As String)
    Dim template_signature As String
    
    If m_code_template_buffer Is Nothing Then Set m_code_template_buffer = New Collection
    template_signature = ExtraVBA.hash12(code_template)
    'detect template duplicates
    If Not ExtraVBA.collection_has_key(template_signature, m_code_template_buffer) Then
        m_code_template_buffer.Add code_template, template_signature
    End If
End Sub

Public Sub print_err_buffer(Optional silent)
    
    If IsMissing(silent) Then silent = False
    If m_err_buffer <> vbNullString Then
        If silent Then
            ConsoleOutput = ConsoleOutput & vbLf & m_err_buffer & vbLf
        Else
            Debug.Print vbLf & m_err_buffer
        End If
    End If
End Sub

Public Sub report(message_type As String, message_content As String, Optional status, Optional err_msg, Optional silent)

    Dim msg_package As Collection
    
    If Not Report_Formatter Is Nothing Then
        Set msg_package = New Collection
        If IsMissing(status) Then
            msg_package.Add vbNullString, REPORT_MSG_STATUS
        Else
            msg_package.Add status, REPORT_MSG_STATUS
        End If
        If IsMissing(err_msg) Then
            msg_package.Add vbNullString, REPORT_MSG_ERR
        Else
            msg_package.Add err_msg, REPORT_MSG_ERR
        End If
        Set msg_package = build_msg_package(message_type, message_content, status, err_msg)
        Report_Formatter.process_msg msg_package, silent:=silent
        Set msg_package = Nothing
    End If
End Sub

Public Function build_msg_package(message_type As String, message_content As String, Optional status, Optional err_msg) As Collection

    Dim msg_package As Collection
    
    If IsMissing(status) Then status = vbNullString
    If IsMissing(err_msg) Then err_msg = vbNullString
    Set msg_package = New Collection
    With msg_package
        .Add message_type, REPORT_MSG_TYPE
        .Add message_content, REPORT_MSG_CONTENT
        .Add status, REPORT_MSG_STATUS
        .Add err_msg, REPORT_MSG_ERR
    End With
    Set build_msg_package = msg_package
End Function

Public Property Get Report_Formatter() As Variant
    If Not IsObject(m_report_formatter) Then
        Set m_report_formatter = New TReportFormatterVerbose
        m_report_formatter.Reporter = Me
    End If
    Set Report_Formatter = m_report_formatter
End Property

Public Property Let Report_Formatter(ByVal new_report_formatter As Variant)
    Set m_report_formatter = new_report_formatter
    m_report_formatter.Reporter = Me
End Property

Public Sub print_code_templates_for_missing_steps(Optional silent)

    Dim code_templates As String
    
    If IsMissing(silent) Then silent = False
    code_templates = get_code_templates_for_missing_steps
    If code_templates <> vbNullString Then
        If silent Then
            ConsoleOutput = vbLf & vbLf & ConsoleOutput & code_templates & vbLf
        Else
            Debug.Print
            Debug.Print
            Debug.Print code_templates
        End If
    End If
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
    End If
    get_code_templates_for_missing_steps = report_output
End Function

Public Property Get ConsoleOutput() As String
    'this is for testing and debugging only, because you can't read the content of the immediate window
    ConsoleOutput = m_console_output
End Property

Public Property Let ConsoleOutput(ByVal new_console_output As String)
    'this is for testing and debugging only, because you can't read the content of the immediate window
    m_console_output = new_console_output
End Property

Public Sub print_report(report_output As String, Optional consecutive, Optional silent)
     
    If IsMissing(silent) Then silent = False
    If IsMissing(consecutive) Then consecutive = False
    If silent Then
        ConsoleOutput = ConsoleOutput & report_output
        If Not consecutive Then ConsoleOutput = ConsoleOutput & vbLf
    ElseIf consecutive Then
        Debug.Print report_output;
    Else
        Debug.Print report_output
    End If
End Sub

Public Sub finish_report(Optional silent)
    
    If IsMissing(silent) Then silent = False
    print_err_buffer silent
    print_code_templates_for_missing_steps silent
End Sub
