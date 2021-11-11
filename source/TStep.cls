VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_step_type As String
Dim m_expressions As Collection
Dim m_parent_clause As Variant
Dim m_elements As Collection
Dim m_docstring As String
Dim m_data_table As TDataTable

Private Sub Class_Initialize()
    m_docstring = vbNullString
End Sub

Private Sub Class_Terminate()
    Set m_expressions = Nothing
    Set m_data_table = Nothing
End Sub

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
            Select Case LCase(TypeName(Me.Expressions(step_element)))
                Case "string"
                    function_name = function_name & "STR"
                Case "long"
                    function_name = function_name & "INT"
                Case "double"
                    function_name = function_name & "DBL"
            End Select
        Else
            function_name = function_name & LCase(step_element)
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
    Dim type_name As String
    
    function_name = vbNullString
    For Each step_element In Me.Elements
        If IsNumeric(step_element) Then
            type_name = LCase(TypeName(Me.Expressions(step_element)))
            If type_name = "long" Then type_name = "integer"
            function_name = function_name & "{" & type_name & "}"
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

Public Function get_step_function_template() As String

    Dim template As String
    
    template = "Public Sub " & Me.get_step_function_name & "(example_context as TContext"
    If Me.Expressions.Count > 0 Then
        template = template & ", step_expressions As Collection"
    End If
    If Not Me.data_table Is Nothing Then
        template = template & ", data_table as TDataTable"
    End If
    template = template & ")" & vbLf
    template = template & Space$(4) & "'" & get_step_comment_name & vbLf
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

Public Property Get parent() As Variant
    Set parent = m_parent_clause
End Property

Public Property Let parent(ByVal parent_clause As Variant)
    Set m_parent_clause = parent_clause
End Property

Public Property Get Elements() As Collection

    If m_elements Is Nothing Then
        Set m_elements = New Collection
    End If
    Set Elements = m_elements
End Property

Public Property Get OriginalStepDefinition() As String
    OriginalStepDefinition = Me.Elements(1) & " " & Me.name
End Property

Public Property Get name() As Variant

    Dim step_name As String
    Dim step_element As Variant
    Dim index As Long
    
    step_name = vbNullString
    For index = 2 To Me.Elements.Count
        step_name = step_name & " "
        step_element = Me.Elements(index)
        If index = Me.Elements.Count And Me.docstring <> vbNullString Then
            step_name = step_name & vbLf & """""""" & vbLf & Me.docstring & vbLf & """"""""
        ElseIf IsNumeric(step_element) Then
            Select Case TypeName(Me.Expressions(step_element))
                Case "String"
                    step_name = step_name & """" & Me.Expressions(step_element) & """"
                Case Else
                    step_name = step_name & CStr(Me.Expressions(step_element))
            End Select
        Else
            step_name = step_name & step_element
        End If
    Next
    If Not Me.data_table Is Nothing Then
        step_name = step_name & vbLf & Me.data_table.get_table_as_text(4)
    End If
    name = Trim(step_name)
End Property

Public Property Get docstring() As String
    docstring = m_docstring
End Property

Public Property Let docstring(ByVal new_docstring As String)
    m_docstring = new_docstring
End Property

Public Property Get data_table() As TDataTable
    Set data_table = m_data_table
End Property

Public Property Set data_table(ByVal new_data_table As TDataTable)
    Set m_data_table = new_data_table
End Property

