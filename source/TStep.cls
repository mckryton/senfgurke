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
Dim m_expressions As Collection
Dim m_parent_clause As Variant
Dim m_elements As Collection
Dim m_docstring As String

Public Property Get SType() As String
    SType = m_step_type
End Property

Public Property Let SType(ByVal step_type As String)
    m_step_type = step_type
End Property

Public Function get_step_function_name() As String
    'vba function names have the same restrictions like vba variables:
    '   https://docs.microsoft.com/de-de/office/vba/language/glossary/vbe-glossary#variable
    '   max len is 255 chars, but compiler crashes on mac os for Office version 16.41 if function name > 63 chars

    Dim function_name As String
    Dim step_element As Variant
    Dim index As Long
    
    function_name = Me.SType
    For index = 2 To Me.Elements.Count
        function_name = function_name & "_"
        step_element = Me.Elements(index)
        If IsNumeric(step_element) Then
            Select Case Me.Expressions(step_element).TypeName
                Case "string"
                    function_name = function_name & "STR"
                Case "integer"
                    function_name = function_name & "INT"
                Case "double"
                    function_name = function_name & "DBL"
            End Select
        Else
            function_name = function_name & step_element
        End If
    Next
    function_name = Replace(function_name, " ", "_")
    function_name = remove_non_ascii_chars(function_name)
    'max len of a function name is 255 chars, but config may require a lower value because of a bug in Office for Mac
    function_name = Left(Trim(function_name), TConfig.MaxStepFunctionNameLength - 13) & "_" & ExtraVBA.hash12(function_name)
    get_step_function_name = function_name
End Function

Private Function get_step_comment_name() As String

    Dim function_name As String
    Dim step_element As Variant
    
    function_name = vbNullString
    For Each step_element In Me.Elements
        If IsNumeric(step_element) Then
            function_name = function_name & "{" & Me.Expressions(step_element).TypeName & "}"
        Else
            function_name = function_name & step_element
        End If
        function_name = function_name & " "
    Next
    'TODO add docstring
    get_step_comment_name = Trim(function_name)
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

Public Sub disassemble_step_name(step_name As String)

    Dim name_element As String
    Dim char_index As Long
    Dim current_char As String
    Dim expression As TStepExpression
    Dim previous_char As String
    
    name_element = vbNullString
    previous_char = " "
    For char_index = 1 To Len(step_name)
        If char_index > 1 Then previous_char = Mid(step_name, char_index - 1, 1)
        current_char = Mid(step_name, char_index, 1)
        If current_char = "\" Then
            'ignore escaped chars, continue parameter search after the escaped char
            char_index = char_index + 2
            name_element = name_element & Mid(step_name, char_index, char_index + 1)
        ElseIf current_char = """" Then
            Set expression = find_string_expression(step_name, char_index)
        ElseIf (IsNumeric(current_char) Or current_char = ".") And previous_char = " " Then
            Set expression = find_numeric_expression(step_name, char_index)
        End If
        If expression Is Nothing Then
            name_element = name_element & current_char
        Else
            Me.Elements.Add Trim(name_element)
            name_element = vbNullString
            'add only a reference to the expression in the list of step name elements
            Me.Elements.Add Me.Expressions.Count + 1
            Me.Expressions.Add expression
            char_index = expression.IndexEnd
            Set expression = Nothing
        End If
    Next
    If Len(name_element) > 0 Then Me.Elements.Add Trim(name_element)
End Sub

Private Function find_numeric_expression(step_name As String, start_index As Long) As TStepExpression

    Dim expression As TStepExpression
    Dim search_index As Long
    Dim current_char As String
    Dim expression_name As String
    
    Set expression = New TStepExpression
    expression.TypeName = "integer"
    expression_name = vbNullString
    search_index = start_index
    expression.IndexStart = start_index
    Do
        current_char = Mid(step_name, search_index, 1)
        If current_char = "." Then expression.TypeName = "double"
        expression_name = expression_name & current_char
        search_index = search_index + 1
    Loop While (IsNumeric(current_char) Or current_char = ".") And search_index <= Len(step_name)
    If current_char = " " Or (IsNumeric(current_char) And search_index > Len(step_name)) Then
        If expression.TypeName = "integer" Then
            expression.Value = CInt(Trim(expression_name))
        Else
            expression.Value = CDbl(Trim(expression_name))
        End If
        If current_char = " " Then search_index = search_index - 1
        expression.IndexEnd = search_index
        Set find_numeric_expression = expression
    Else
        'numeric value ends on a letter e.g. 2nd
        Set find_numeric_expression = Nothing
    End If
End Function

Private Function find_string_expression(step_name As String, start_index As Long) As TStepExpression

    Dim expression As TStepExpression
    Dim found_matching_quote As Boolean
    Dim search_index As Long
    Dim matching_position As Long

    Set expression = New TStepExpression
    expression.TypeName = "string"
    'found string parameter, look for matching quotes
    found_matching_quote = False
    search_index = start_index
    expression.IndexStart = start_index
    Do
        'ignore escaped quotes as matching quotes
        matching_position = InStr(search_index + 1, step_name, """")
        If Mid(step_name, matching_position - 1, 1) = "\" Then
            search_index = matching_position
        Else
            found_matching_quote = True
        End If
    Loop While found_matching_quote = False And matching_position > 0
    If matching_position = 0 Then matching_position = Len(step_name)
    expression.Value = Mid(step_name, start_index + 1, matching_position - start_index - 1)
    expression.Value = Replace(expression.Value, "\""", """")
    expression.IndexEnd = matching_position
    Set find_string_expression = expression
End Function

Public Function get_step_function_template() As String

    Dim template As String
    
    template = "Public Sub " & Me.get_step_function_name & "("
    If Me.Expressions.Count > 0 Then
        template = template & "step_expressions As Collection"
    End If
    template = template & ")" & vbLf
    template = template & "    '" & get_step_comment_name & vbLf
    template = template & Space$(4) & "pending" & vbLf
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

Public Property Get Parent() As Variant
    Set Parent = m_parent_clause
End Property

Public Property Let Parent(ByVal parent_clause As Variant)
    Set m_parent_clause = parent_clause
End Property

Public Sub parse_step_definition(step_definition As String)
    
    Dim step_head As String
    Dim step_name As String
    
    step_head = Split(Trim(step_definition), " ")(0)
    step_name = Trim(Right(step_definition, Len(step_definition) - Len(step_head) - 1))
    Me.Elements.Add step_head
    If InStr("Given When Then", step_head) > 0 And Len(step_head) > 3 Then
        Me.SType = step_head
    Else
        'assume that this step is not yet added to it's parent clause
        If Me.Parent.Steps.Count = 0 Then
            Me.SType = STEP_TYPE_GIVEN
        Else
            Me.SType = Me.Parent.Steps(Me.Parent.Steps.Count).SType
        End If
    End If
    disassemble_step_name step_name
End Sub

Public Property Get Elements() As Collection

    If m_elements Is Nothing Then
        Set m_elements = New Collection
    End If
    Set Elements = m_elements
End Property

Public Property Get OriginalStepDefinition() As String
    OriginalStepDefinition = Me.Elements(1) & " " & Me.Name
End Property

Public Property Get Name() As Variant

    Dim step_name As String
    Dim step_element As Variant
    Dim index As Long
    
    step_name = vbNullString
    For index = 2 To Me.Elements.Count
        step_name = step_name & " "
        step_element = Me.Elements(index)
        If IsNumeric(step_element) Then
            Select Case Me.Expressions(step_element).TypeName
                Case "string"
                    step_name = step_name & """" & Me.Expressions(step_element).Value & """"
                Case Else
                    step_name = step_name & CStr(Me.Expressions(step_element).Value)
            End Select
        Else
            step_name = step_name & step_element
        End If
    Next
    Name = Trim(step_name)
End Property

Public Property Get Docstring() As String
    Docstring = m_docstring
End Property

Public Property Let Docstring(ByVal new_docstring As String)
    m_docstring = new_docstring
End Property
