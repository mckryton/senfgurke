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
            Case LINE_TYPE_TABLE_ROW
                If new_step_list.Count = 0 Then
                    TError.raise ERR_ID_STEP_SYNTAX_TABLE_WITHOUT_STEP, _
                        "TStepParser.parse_step_list"
                End If
                parent_feature.parsed_lines = parent_feature.parsed_lines - 1
                Set data_table = parse_table(feature_text, parent_feature)
                Set current_step = new_step_list(new_step_list.Count)
                Set current_step.data_table = data_table
            Case LINE_TYPE_EXAMPLE_START, LINE_TYPE_RULE_START, LINE_TYPE_TAGS, LINE_TYPE_OUTLINE_START
                'new rules, examples or tags will complete the step list
                Set parse_step_list = new_step_list
                parent_feature.parsed_lines = parent_feature.parsed_lines - 1
                Exit Function
            Case Else
                'ignore empty lines
                If Not Trim(line) = "" Then
                    TError.raise ERR_ID_FEATURE_SYNTAX_UNEXPECTED_LINE, _
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
            TError.raise ERR_ID_STEP_SYNTAX_TABLE_DUPLICATE_COLUMN, _
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
    Dim step As TStep
    
    Set step = New TStep
    step_line = Trim(step_line)
    step.OriginalStepLine = step_line
    step_head = Split(step_line, " ")(0)
    'step head is the first word in a step line e.g. Given in 'Given a step' or And in 'And another step'
    step.Elements.Add step_head
    'step is only one of those: Given, When or Then
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
    Set parse_step_line = step
End Function

