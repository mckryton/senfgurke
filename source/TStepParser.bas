Attribute VB_Name = "TStepParser"
Option Explicit

Public Function parse_step_list(feature_text As String, parent_feature As TFeature) As Collection

    Dim line As String
    Dim current_step As TStep
    Dim feature_lines As Variant
    Dim new_step_list As Collection
    Dim data_table As TDataTable
    
    Set new_step_list = New Collection
    feature_lines = Split(Trim(feature_text), vbLf)
    Do While parent_feature.parsed_lines < UBound(feature_lines)
        parent_feature.parsed_lines = parent_feature.parsed_lines + 1
        line = CStr(feature_lines(parent_feature.parsed_lines))
        Select Case get_line_type(line)
            Case LINE_TYPE_COMMENT
                'ignore comments
            Case LINE_TYPE_STEP
                new_step_list.Add parse_step_line(line, new_step_list)
            Case LINE_TYPE_DOCSTRING_LIMIT
                Set current_step = new_step_list(new_step_list.Count)
                current_step.docstring = parse_docstring(feature_text, parent_feature)
                current_step.Expressions.Add current_step.docstring
                current_step.Elements.Add current_step.Expressions.Count
            Case LINE_TYPE_TABLE_ROW
                If new_step_list.Count = 0 Then
                    TError.raise_pre_defined_error ERR_ID_STEP_SYNTAX_TABLE_WITHOUT_STEP, _
                        "TStepParser.parse_step_list"
                End If
                parent_feature.parsed_lines = parent_feature.parsed_lines - 1
                Set data_table = parse_table(feature_text, parent_feature)
                Set current_step = new_step_list(new_step_list.Count)
                Set current_step.data_table = data_table
            Case LINE_TYPE_EXAMPLE_START, LINE_TYPE_RULE_START, LINE_TYPE_TAGS
                'new rules, examples or tags will complete the step list
                Set parse_step_list = new_step_list
                parent_feature.parsed_lines = parent_feature.parsed_lines - 1
                Exit Function
            Case Else
                'ignore empty lines
                If Not Trim(line) = "" Then
                    TError.raise_pre_defined_error ERR_ID_FEATURE_SYNTAX_UNEXPECTED_LINE, _
                        "TStepParser.parse_step_list", _
                        Array(parent_feature.parsed_lines, line)
                End If
        End Select
    Loop
    Set parse_step_list = new_step_list
End Function

Public Function parse_docstring(feature_text As String, parent_feature As TFeature) As String

    Dim line As String
    Dim docstring As String
    Dim feature_lines As Variant
    
    docstring = vbNullString
    feature_lines = Split(Trim(feature_text), vbLf)
    Do While parent_feature.parsed_lines < UBound(feature_lines)
        parent_feature.parsed_lines = parent_feature.parsed_lines + 1
        line = CStr(feature_lines(parent_feature.parsed_lines))
        If Trim(line) = """""""" Then
            'remove the last linebreak
            docstring = Left(docstring, Len(docstring) - 1)
            parse_docstring = align_textblock(docstring)
            Exit Function
        Else
            docstring = docstring & line & vbLf
        End If
    Loop
    'remove the last linebreak
    docstring = Left(docstring, Len(docstring) - 1)
    parse_docstring = align_textblock(docstring)
End Function

Public Function parse_table(feature_text As String, parent_feature As TFeature) As TDataTable

    Dim line As String
    Dim data_table As TDataTable
    Dim feature_lines As Variant

    Set data_table = Nothing
    feature_lines = Split(Trim(feature_text), vbLf)
    Do While parent_feature.parsed_lines < UBound(feature_lines)
        parent_feature.parsed_lines = parent_feature.parsed_lines + 1
        line = CStr(feature_lines(parent_feature.parsed_lines))
        Select Case TFeatureParser.get_line_type(line)
            Case LINE_TYPE_COMMENT
                'ignore comments
            Case LINE_TYPE_TABLE_ROW
                If data_table Is Nothing Then
                    Set data_table = create_data_table(line)
                Else
                    data_table.add_row line
                End If
            Case Else
                Set parse_table = data_table
                parent_feature.parsed_lines = parent_feature.parsed_lines - 1
            Exit Function
        End Select
    Loop
    Set parse_table = data_table
End Function

Private Function create_data_table(table_header_line As String) As TDataTable

    Dim data_table As TDataTable
    Dim column_name As String
    Dim column_names As Variant
    Dim column_index As Integer
    Dim was_added As Boolean

    table_header_line = Trim(table_header_line)
    Set data_table = New TDataTable
    column_names = Split(table_header_line, "|")
    For column_index = 1 To UBound(column_names) - 1
        column_name = Trim(CStr(column_names(column_index)))
        was_added = data_table.add_column_name(column_name)
        If Not was_added Then
            TError.raise_pre_defined_error ERR_ID_STEP_SYNTAX_TABLE_DUPLICATE_COLUMN, _
                "TStepParser.create_data_table", _
                Array(column_name)
        End If
    Next
    Set create_data_table = data_table
End Function

Public Function align_textblock(indented_text As String) As String
    
    Dim trimmed_text As String
    Dim min_indention As Integer
    Dim lines As Variant
    Dim line As Variant
    Dim indention As Long

    min_indention = -1
    lines = Split(indented_text, vbLf)
    For Each line In lines
        If Trim(line) <> vbNullString Then
            indention = Len(line) - Len(LTrim(line))
            If min_indention = -1 Or indention < min_indention Then min_indention = indention
        End If
    Next
    If min_indention = -1 Then
        trimmed_text = indented_text
    Else
        For Each line In lines
            If Len(line) > min_indention Then
                trimmed_text = trimmed_text & Right(line, Len(line) - min_indention) & vbLf
            Else
                trimmed_text = trimmed_text & line & vbLf
            End If
        Next
        'remove the last linebreak
        If Len(trimmed_text) > 0 Then trimmed_text = Left(trimmed_text, Len(trimmed_text) - 1)
    End If
    align_textblock = trimmed_text
End Function

Public Function parse_step_line(step_line As String, step_list As Collection) As TStep
    
    Dim step_head As String
    Dim step_name As String
    Dim step As TStep
    
    step_line = Trim(step_line)
    Set step = New TStep
    step_head = Split(step_line, " ")(0)
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
            char_index = char_index + 1
            If char_index <= Len(step_name) Then
                current_char = Mid(step_name, char_index, 1)
                name_element = name_element & current_char
            End If
        Else
            If current_char = """" Then
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

