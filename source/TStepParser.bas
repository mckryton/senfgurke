Attribute VB_Name = "TStepParser"
Option Explicit

Public Function parse_steps_new(feature_text As String, step_list_start_index As Long, step_list As Collection) As Long

    Dim line_index As Long
    Dim line As String
    Dim is_docstring As Boolean
    Dim docstring_value As String
    Dim current_step As TStep
    Dim feature_lines As Variant
    
    feature_lines = Split(Trim(feature_text), vbLf)
    is_docstring = False
    docstring_value = vbNullString
    For line_index = step_list_start_index To UBound(feature_lines)
        'this normalizes docstrings by triming every line -> this way docstrings can be compared by content
        '  but not by indention!
        line = Trim(feature_lines(line_index))
        If Right(line, 3) = """""""" And is_docstring Then
            is_docstring = False
            Set current_step = step_list(step_list.Count)
            current_step.Docstring = Right(docstring_value, Len(docstring_value) - 1)
            current_step.Expressions.Add current_step.Docstring
            current_step.Elements.Add current_step.Expressions.Count
            docstring_value = vbNullString
        ElseIf is_docstring Then
            If docstring_value = vbNullString Then
                ' mark the start of the docstring with a # so that leading empty lines are recognized
                docstring_value = "#" & line
            Else
                docstring_value = docstring_value & vbLf & line
            End If
        ElseIf Left(line, 3) = """""""" And Not is_docstring Then
            is_docstring = True
        ElseIf TFeatureParser.is_step_line(line) Then
            step_list.Add parse_step_line(line, step_list)
        ElseIf TFeatureParser.is_section_definition_line(line) Or Trim(line) = vbNullString Or TFeatureParser.is_tag_line(line) Then
            'example is finished either with next example, empty line or tag line
            parse_steps_new = line_index - 1
            Exit Function
        End If
    Next
    parse_steps_new = line_index
End Function

Public Function parse_step_line(step_line As String, step_list As Collection) As TStep
    
    Dim step_head As String
    Dim step_name As String
    Dim step As TStep
    
    Set step = New TStep
    step_head = Split(Trim(step_line), " ")(0)
    step_name = Trim(Right(step_line, Len(step_line) - Len(step_head) - 1))
    step.Elements.Add step_head
    If InStr("Given When Then", step_head) > 0 And Len(step_head) > 3 Then
        step.SType = step_head
    Else
        'assume that this step is not yet added to it's parent clause
        If step_list.Count = 0 Then
            step.SType = STEP_TYPE_GIVEN
        Else
            step.SType = step_list(step_list.Count).SType
        End If
    End If
    find_step_expressions step_name, step
    Set parse_step_line = step
End Function

Private Sub find_step_expressions(step_name As String, new_step As TStep)

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
        ElseIf (IsNumeric(current_char) Or current_char = ".") And InStr(" ,([{", previous_char) > 0 Then
            Set expression = find_numeric_expression(step_name, char_index)
        End If
        If expression Is Nothing Then
            name_element = name_element & current_char
        Else
            new_step.Elements.Add Trim(name_element)
            name_element = vbNullString
            'add only a reference to the expression in the list of step name elements
            new_step.Elements.Add new_step.Expressions.Count + 1
            new_step.Expressions.Add expression.value
            char_index = expression.IndexEnd
            Set expression = Nothing
        End If
    Next
    If Len(name_element) > 0 Then new_step.Elements.Add Trim(name_element)
End Sub

Private Function find_numeric_expression(step_name As String, start_index As Long) As TStepExpression

    Dim expression As TStepExpression
    Dim search_index As Long
    Dim current_char As String
    Dim expression_name As String
    Dim expression_is_valid As Boolean
    
    Set expression = New TStepExpression
    expression.TypeName = "integer"
    expression_name = vbNullString
    search_index = start_index
    expression_is_valid = True
    expression.IndexStart = start_index
    Do
        current_char = Mid(step_name, search_index, 1)
        If current_char = "." Then
            If expression.TypeName = "double" Then
                'a numeric expression with more than one decimal separator is probably a date
                expression_is_valid = False
            Else
                expression.TypeName = "double"
            End If
        End If
        expression_name = expression_name & current_char
        search_index = search_index + 1
    Loop While (IsNumeric(current_char) Or current_char = ".") And search_index <= Len(step_name)
    If (InStr(" ,}])", current_char) > 0 Or (IsNumeric(current_char)) And search_index > Len(step_name)) And expression_is_valid Then
        If expression.TypeName = "integer" Then
            If Not IsNumeric(current_char) Then expression_name = Left(expression_name, Len(expression_name) - 1)
            expression.value = CLng(expression_name)
        Else
            expression.value = CDbl(Trim(expression_name))
        End If
        If InStr(" ,}])", current_char) > 0 Then search_index = search_index - 1
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
    expression.value = Mid(step_name, start_index + 1, matching_position - start_index - 1)
    expression.value = Replace(expression.value, "\""", """")
    expression.IndexEnd = matching_position
    Set find_string_expression = expression
End Function

