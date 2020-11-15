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
Dim m_expressions As Collection

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
    Me.Expressions = extract_step_expressions(function_name)
    function_name = Me.Expressions("function_name_with_placeholders")
    Me.Expressions.remove "function_name_with_placeholders"
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

Public Function extract_step_expressions(function_name As String) As Collection

    Dim step_expressions As Collection
    Dim function_name_with_placeholders As String
    Dim char_index As Long
    Dim current_char As String
    Dim next_expression As Collection
    Dim previous_char As String
    
    function_name_with_placeholders = vbNullString
    Set step_expressions = New Collection
    For char_index = 1 To Len(function_name)
        If char_index > 1 Then previous_char = Mid(function_name, char_index - 1, 1) Else previous_char = " "
        current_char = Mid(function_name, char_index, 1)
        If current_char = "\" Then
            'ignore escaped chars, continue parameter search after the escaped char
            char_index = char_index + 2
            function_name_with_placeholders = function_name_with_placeholders & Mid(function_name, char_index, char_index + 1)
        ElseIf current_char = """" Then
            Set next_expression = find_string_expression(function_name, char_index)
        ElseIf (IsNumeric(current_char) Or current_char = ".") And previous_char = " " Then
            Set next_expression = find_numeric_expression(function_name, char_index)
        End If
        If next_expression Is Nothing Then
            function_name_with_placeholders = function_name_with_placeholders & current_char
        Else
            add_expression next_expression, step_expressions
            function_name_with_placeholders = function_name_with_placeholders & next_expression("type")
            char_index = next_expression("end_index")
            Set next_expression = Nothing
        End If
    Next
    step_expressions.Add function_name_with_placeholders, "function_name_with_placeholders"
    Set extract_step_expressions = step_expressions
End Function

Private Function find_numeric_expression(function_name As String, start_index As Long) As Collection

    Dim expression As Collection
    Dim search_index As Long
    Dim current_char As String
    Dim expression_type As String
    Dim new_expression As String
    
    expression_type = "INT"
    new_expression = vbNullString
    search_index = start_index
    Do
        current_char = Mid(function_name, search_index, 1)
        If current_char = "." Then expression_type = "DBL"
        new_expression = new_expression & current_char
        search_index = search_index + 1
    Loop While (IsNumeric(current_char) Or current_char = ".") And search_index <= Len(function_name)
    If current_char = " " Or (IsNumeric(current_char) And search_index > Len(function_name)) Then
        Set expression = New Collection
        expression.Add Trim(new_expression), "value"
        expression.Add expression_type, "type"
        If current_char = " " Then search_index = search_index - 2
        expression.Add search_index, "end_index"
    End If
    Set find_numeric_expression = expression
End Function

Private Function find_string_expression(function_name As String, start_index As Long) As Collection

    Dim expression As Collection
    Dim found_matching_quote As Boolean
    Dim search_index As Long
    Dim matching_position As Long

    Set expression = New Collection
    'found string parameter, look for matching quotes
    found_matching_quote = False
    search_index = start_index
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
    expression.Add Mid(function_name, start_index + 1, matching_position - start_index - 1), "value"
    expression.Add "STR", "type"
    expression.Add matching_position, "end_index"
    Set find_string_expression = expression
End Function

Private Sub add_expression(expression As Collection, extracted_expressions As Collection)

    Select Case expression("type")
        Case "INT"
            extracted_expressions.Add CInt(expression("value"))
        Case "DBL"
            extracted_expressions.Add CDbl(expression("value"))
        Case Else
            extracted_expressions.Add CStr(Replace(expression("value"), "\", ""))
    End Select
End Sub

Public Function get_step_function_template() As String

    Dim template As String
    
    template = "Public Sub " & Me.get_step_function_name & "("
    If Me.Expressions.Count > 0 Then
        template = template & "step_expressions As Collection"
    End If
    template = template & ")" & vbLf
    template = template & "    '" & Me.StepDefinition & vbLf & vbLf
    template = template & "End Sub"
    get_step_function_template = template
End Function

Public Property Get Expressions() As Collection
    If m_expressions Is Nothing Then
        Set m_expressions = New Collection
    End If
    Set Expressions = m_expressions
End Property

Public Property Let Expressions(ByVal new_expressions As Collection)
    Set m_expressions = new_expressions
End Property

Private Sub Class_Terminate()
    Set m_expressions = Nothing
End Sub
