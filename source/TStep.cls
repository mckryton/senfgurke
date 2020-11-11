VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_step_type As String
Dim m_step_head As String
Dim m_step_name As String
Dim m_parameters As Collection

Public Property Get SType() As String
    SType = m_step_type
End Property

Public Property Let SType(ByVal step_type As String)
    m_step_type = step_type
End Property

Public Property Get Head() As String
    Head = m_step_head
End Property

Public Property Let Head(ByVal step_head As String)
    m_step_head = step_head
End Property

Public Property Get Name() As String
    Name = m_step_name
End Property

Public Property Let Name(ByVal step_name As String)
    m_step_name = Trim(step_name)
End Property

Public Property Get OriginalStepDefinition() As String
    OriginalStepDefinition = Me.Head & " " & Me.Name
End Property

Public Property Get StepDefinition() As String
    StepDefinition = Me.SType & " " & Me.Name
End Property

Public Function get_step_function_name() As String
    'vba function names have the same restrictions like vba variables:
    '   https://docs.microsoft.com/de-de/office/vba/language/glossary/vbe-glossary#variable
    '   max len is 255 chars, but compiler crashes on mac os for Office version 16.41 if function name > 63 chars

    Dim current_char As String
    Dim function_name As String
    Dim index As Integer
    
    function_name = Me.StepDefinition
    'function_name = Replace(function_name, "\""", "_")
    Me.parameters = extract_step_parameters(function_name)
    function_name = Me.parameters("function_name_with_placeholders")
    Me.parameters.remove "function_name_with_placeholders"
    function_name = Replace(function_name, " ", "_")
    function_name = remove_non_ascii_chars(function_name)
    'max len of a function name is 255 chars, but config may require a lower value because of a bug in Office for Mac
    function_name = Left(Trim(function_name), TConfig.MaxStepFunctionNameLength - 13) & "_" & ExtraVBA.hash12(function_name)
    get_step_function_name = function_name
End Function

Private Function remove_non_ascii_chars(function_name As String) As String

    Dim current_char As String
    Dim index As Integer
    
    index = 1
    Do While index <= Len(function_name)
        current_char = Mid(function_name, index, 1)
        'a vba function name may contain any ascii letter, numbers and underscore
        If Not ( _
                 (Asc(current_char) > 47 And Asc(current_char) < 58) _
                Or (Asc(current_char) > 64 And Asc(current_char) < 91) _
                Or Asc(current_char) = 95 _
                Or (Asc(current_char) > 96 And Asc(current_char) < 123) _
                ) Then
            function_name = Replace(function_name, current_char, "")
        Else
            index = index + 1
        End If
    Loop
    remove_non_ascii_chars = function_name
End Function

Public Function extract_step_parameters(function_name As String) As Collection

    Dim step_parameters As Collection
    Dim function_name_with_placeholders As String
    Dim char_index As Long
    Dim current_char As String
    Dim param_value As String
    Dim matching_position As Long
    Dim param_type As String
    Dim found_matching_quote As Boolean
    Dim search_index As Long
    
    function_name_with_placeholders = vbNullString
    param_type = vbNullString
    Set step_parameters = New Collection
    param_value = vbNullString
    For char_index = 1 To Len(function_name)
        current_char = Mid(function_name, char_index, 1)
        If current_char = "\" Then
            'ignore escaped chars, continue parameter search after the escaped char
            char_index = char_index + 2
            function_name_with_placeholders = function_name_with_placeholders & Mid(function_name, char_index, char_index + 1)
        ElseIf current_char = """" Then
            'found string parameter, look for matching quotes
            found_matching_quote = False
            search_index = char_index
            Do
                'ignore escaped quotes as matching quotes
                matching_position = InStr(search_index + 1, function_name, """")
                If Mid(function_name, matching_position - 1, 1) = "\" Then
                    search_index = matching_position
                Else
                    found_matching_quote = True
                End If
            Loop While found_matching_quote = False And matching_position > 0
            If matching_position = 0 Then matching_position = Len(function_name)
            param_value = Mid(function_name, char_index + 1, matching_position - char_index - 1)
            param_type = "STR"
            add_parameter param_value, param_type, step_parameters
            function_name_with_placeholders = function_name_with_placeholders & param_type
            char_index = matching_position
            param_value = vbNullString
            param_type = vbNullString
        ElseIf IsNumeric(current_char) Or (param_value <> vbNullString And current_char = ".") Then
            'add number or separator to the numeric parameter
            param_value = param_value & current_char
            If current_char = "." Then param_type = "DBL"
        ElseIf param_value <> vbNullString Then
            'reached the end of a numeric parameter
            add_parameter param_value, param_type, step_parameters
            function_name_with_placeholders = function_name_with_placeholders & param_type & current_char
            param_value = vbNullString
            param_type = vbNullString
        Else
            function_name_with_placeholders = function_name_with_placeholders & current_char
        End If
    Next
    If param_value <> vbNullString Then
        add_parameter param_value, param_type, step_parameters
        function_name_with_placeholders = function_name_with_placeholders & param_type
    End If
    step_parameters.Add function_name_with_placeholders, "function_name_with_placeholders"
    Set extract_step_parameters = step_parameters
End Function

Private Sub add_parameter(param_value As String, param_type As String, parameters As Collection)

    If param_type = vbNullString Then param_type = "INT"
    Select Case param_type
        Case "INT"
            parameters.Add CInt(param_value)
        Case "DBL"
            parameters.Add CDbl(param_value)
        Case Else
            param_value = Replace(param_value, "\", "")
            parameters.Add CStr(param_value)
    End Select
End Sub

Public Function get_step_function_template() As String

    Dim template As String
    
    template = "Public Sub " & Me.get_step_function_name & "("
    If Me.parameters.Count > 0 Then
        template = template & "step_parameters As Collection"
    End If
    template = template & ")" & vbLf
    template = template & "    '" & Me.StepDefinition & vbLf & vbLf
    template = template & "End Sub"
    get_step_function_template = template
End Function

Public Property Get parameters() As Collection
    If m_parameters Is Nothing Then
        Set m_parameters = New Collection
    End If
    Set parameters = m_parameters
End Property

Public Property Let parameters(ByVal new_parameters As Collection)
    Set m_parameters = new_parameters
End Property

Private Sub Class_Terminate()
    Set m_parameters = Nothing
End Sub
