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
    function_name = Replace(function_name, " ", "_")
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
    'max len of a function name is 255 chars
    function_name = Left(Trim(function_name), TConfig.MaxStepFunctionNameLength - 13) & "_" & ExtraVBA.hash12(Me.StepDefinition)
    get_step_function_name = function_name
End Function

Public Function get_step_function_template() As String

    Dim template As String
    
    template = "Public Sub " & Me.get_step_function_name & "()" & vbLf
    template = template & "    '" & Me.StepDefinition & vbLf & vbLf
    template = template & "End Sub"
    get_step_function_template = template
End Function
