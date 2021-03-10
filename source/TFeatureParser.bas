Attribute VB_Name = "TFeatureParser"
Option Explicit

Const KEYWORDS = "Example,Scenario,Scenario Outline,Rule,Ability,Business Needs,Feature,Background"

Public Function parse_feature(gherkin_text As String) As TFeature

    Dim parsed_feature As TFeature
    Dim lines As Variant
    Dim line As String
    Dim keyword_value As Variant
    Dim current_clause As Variant
    Dim step_head_name As Variant
    Dim has_example_steps As Boolean
    Dim line_index As Long
    Dim clause_tags As Collection
    
    Set parsed_feature = parse_feature_definition(gherkin_text)
    Set current_clause = Nothing
    lines = Split(Trim(gherkin_text), vbLf)
    Set clause_tags = clone_tags(parsed_feature.tags)
    If parsed_feature.ErrorStatus = vbNullString Then
        For line_index = parsed_feature.ParsedLinesIndex + 1 To UBound(lines)
            line = Trim(lines(line_index))
            If is_clause_definition_line(line) Then
                keyword_value = read_keyword_value(CStr(line))
                If Not current_clause Is Nothing Then trim_description_linebreaks current_clause
                Set current_clause = update_feature(parsed_feature, CStr(keyword_value(0)), CStr(keyword_value(1)), clause_tags, line)
                If CStr(keyword_value(0)) = CLAUSE_TYPE_EXAMPLE Or CStr(keyword_value(0)) = CLAUSE_TYPE_BACKGROUND Then
                    line_index = parse_steps(lines, line_index, current_clause)
                End If
                Set clause_tags = clone_tags(parsed_feature.tags)
            ElseIf is_tag_line(line) Then
                add_tags line, clause_tags
            ElseIf is_comment_line(line) Then
                'ignore comments
            Else
                If Not TypeName(current_clause) = "TExample" Or TypeName(current_clause) = "TBackground" Then
                    add_description CStr(line), parsed_feature
                End If
            End If
        Next
    End If
    Set parse_feature = parsed_feature
End Function

Private Sub trim_description_linebreaks(feature_clause)

    Do While Right(feature_clause.Description, 1) = vbLf
        feature_clause.Description = Left(feature_clause.Description, Len(feature_clause.Description) - 1)
    Loop
    Do While Left(feature_clause.Description, 1) = vbLf
        feature_clause.Description = Right(feature_clause.Description, Len(feature_clause.Description) - 1)
    Loop
End Sub

Public Function parse_steps(feature_lines As Variant, example_start_index As Long, current_clause As Variant) As Long

    Dim line_index As Long
    Dim line As String
    Dim is_docstring As Boolean
    Dim docstring_value As String
    Dim current_step As TStep
    
    is_docstring = False
    docstring_value = vbNullString
    For line_index = example_start_index + 1 To UBound(feature_lines)
        'this normalizes dostrings by triming every line -> this way docstrings can be compared by content
        '  but not by indention!
        line = Trim(feature_lines(line_index))
        If Right(line, 3) = """""""" And is_docstring Then
            is_docstring = False
            Set current_step = current_clause.Steps(current_clause.Steps.Count)
            current_step.Docstring = Right(docstring_value, Len(docstring_value) - 1)
            current_step.Expressions.Add docstring_value
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
        ElseIf is_step_line(line) Then
            current_clause.Steps.Add create_step(line, current_clause)
        ElseIf is_clause_definition_line(line) Or Trim(line) = "" Or is_tag_line(line) Then
            'example is finished either with next example, empty line or tag line
            parse_steps = line_index - 1
            Exit Function
        End If
    Next
    parse_steps = line_index
End Function

Private Function parse_feature_definition(gherkin_text As String) As TFeature
    
    Dim parsed_feature As TFeature
    Dim line_index As Long
    Dim feature_lines As Variant
    Dim header_finished As Boolean
    Dim feature_spec As Variant
    Dim line As String

    Set parsed_feature = New TFeature
    header_finished = False
    line_index = 0
    feature_lines = Split(gherkin_text, vbLf)
    Do While line_index <= UBound(feature_lines) And header_finished = False
        line = CStr(feature_lines(line_index))
        If is_tag_line(line) Then
            parsed_feature.add_tags line
            line_index = line_index + 1
        ElseIf is_comment_line(line) Then
            line_index = line_index + 1
        Else
            feature_spec = read_keyword_value(line)
            If Not CStr(feature_spec(0)) = CLAUSE_TYPE_FEATURE Then
                parsed_feature.ErrorStatus = "Feature lacks feature keyword at the beginning"
            Else
                parsed_feature.Name = CStr(feature_spec(1))
            End If
            header_finished = True
        End If
    Loop
    parsed_feature.ParsedLinesIndex = line_index
    Set parse_feature_definition = parsed_feature
End Function

Private Function is_clause_definition_line(feature_line As String) As Boolean
    
    Dim keyword  As Variant
    Dim clause_name As String
    
    'clause definition format is /^<keyword>:[^:]*$/
    If UBound(Split(Trim(feature_line), ":")) > 0 Then
        clause_name = Split(Trim(feature_line), ":")(0)
        For Each keyword In Split(KEYWORDS, ",")
            If clause_name = CStr(keyword) Then
               is_clause_definition_line = True
               Exit Function
            End If
        Next
    End If
    is_clause_definition_line = False
End Function

Private Function is_comment_line(feature_line As String) As Boolean

    If Left(Trim(feature_line), 1) = "#" Then
        is_comment_line = True
    Else
        is_comment_line = False
    End If
End Function

Private Function is_tag_line(feature_line As String) As Boolean

    If Left(Trim(feature_line), 1) = "@" Then
        is_tag_line = True
    Else
        is_tag_line = False
    End If
End Function

Private Function is_step_line(feature_line As String) As Boolean

    Dim first_word As String
    
    first_word = Split(Trim(feature_line) & " ", " ")(0)
    If InStr("Given When Then And But", first_word) > 0 And Len(first_word) > 2 Then
        is_step_line = True
    Else
        is_step_line = False
    End If
End Function

Public Sub add_tags(feature_line As String, tags As Collection)

    Dim tag_list As Variant
    Dim index As Long
    Dim tag As Variant
    
    tag_list = Split(Trim(feature_line), " ")
    For Each tag In tag_list
        If Len(tag) > 1 Then
            If Left(tag, 1) = "@" Then
                If Not ExtraVBA.collection_has_key(tag, tags) Then
                    tags.Add tag, tag
                End If
            End If
        End If
    Next
End Sub

Private Function update_feature(parsed_feature As TFeature, clause_keyword As String, clause_name As String, clause_tags As Collection, feature_line As String) As Variant

    Dim current_clause As Variant
    
    Select Case clause_keyword
        Case CLAUSE_TYPE_RULE
            Set current_clause = create_rule(rule_name:=clause_name)
            parsed_feature.Clauses.Add current_clause
        Case CLAUSE_TYPE_EXAMPLE
            Set current_clause = create_example(example_head:=CStr(clause_keyword), example_name:=clause_name, example_tags:=clause_tags, feature_line:=feature_line)
            parsed_feature.Clauses.Add current_clause
        Case CLAUSE_TYPE_BACKGROUND
            Set current_clause = parsed_feature.Background
        Case Else
            Debug.Print "PARSE ERROR: unknown clause >" & clause_keyword & "<"
    End Select
    Set update_feature = current_clause
End Function

Private Function read_keyword_value(text_line As String) As Variant

    Dim line_items As Variant
    Dim keyword As String
    
    line_items = Split(Trim(text_line), ":")
    keyword = vbNullString
    If UBound(line_items) > 0 Then
        Select Case line_items(0)
            Case "Feature"
                keyword = CLAUSE_TYPE_FEATURE
            Case "Ability"
                keyword = CLAUSE_TYPE_FEATURE
            Case "Business Need"
                keyword = CLAUSE_TYPE_FEATURE
            Case "Background"
                keyword = CLAUSE_TYPE_BACKGROUND
            Case "Rule"
                keyword = CLAUSE_TYPE_RULE
            Case "Scenario"
                keyword = CLAUSE_TYPE_EXAMPLE
            Case "Scenario Outline"
                keyword = CLAUSE_TYPE_EXAMPLE
            Case "Example"
                keyword = CLAUSE_TYPE_EXAMPLE
            Case Else
                Debug.Print "PARSE ERROR: unknown keyword >" & line_items(0) & "<"
        End Select
    End If
    read_keyword_value = Array(keyword, Trim(Right(text_line, Len(text_line) - InStr(text_line, ":"))))
End Function

Private Function create_rule(rule_name As String) As TRule

    Dim new_rule As TRule
    
    Set new_rule = New TRule
    new_rule.Name = rule_name
    Set create_rule = new_rule
End Function

Private Function create_example(example_head As String, example_name As String, example_tags As Collection, feature_line As String) As TExample

    Dim new_example As TExample
    
    Set new_example = New TExample
    new_example.Head = example_head
    new_example.Name = example_name
    new_example.tags = example_tags
    new_example.OriginalHeadline = feature_line
    Set create_example = new_example
End Function

Private Sub add_description(line As String, feature As TFeature)

    Dim clause As Variant
    
    If feature.Clauses.Count = 0 Then
        If feature.Description = vbNullString Then
            feature.Description = Trim(line)
        Else
            feature.Description = feature.Description & vbLf & Trim(line)
        End If
    Else
        Set clause = feature.Clauses(feature.Clauses.Count)
        If clause.Description = vbNullString Then
            clause.Description = Trim(line)
        Else
            clause.Description = clause.Description & vbLf & Trim(line)
        End If
    End If
End Sub

Private Function read_step_head_name(line As String) As Variant

    Dim step_type As String
    Dim trim_line As String

    trim_line = Trim(line)
    If Len(trim_line) > 0 Then
        step_type = Split(trim_line, " ")(0)
        If InStr("Given When Then And But", step_type) > 0 Then
            read_step_head_name = Array(step_type, Right(trim_line, Len(trim_line) - InStr(trim_line, " ")))
            Exit Function
        End If
    End If
    read_step_head_name = Array(vbNullString, vbNullString)
End Function

Public Function parse_loaded_features(feature_list As Collection) As Collection
    
    Dim gherkin_text As Variant
    Dim parsed_features As Collection

    Set parsed_features = New Collection
    For Each gherkin_text In feature_list
        parsed_features.Add parse_feature(CStr(gherkin_text))
    Next
    Set parse_loaded_features = parsed_features
End Function

Private Function clone_tags(source_tags As Collection) As Collection

    Dim tag As Variant
    Dim target_tags As Collection
    
    Set target_tags = New Collection
    For Each tag In source_tags
        target_tags.Add tag, tag
    Next
    Set clone_tags = target_tags
End Function

Public Function create_step(step_definition As String, Optional parent_clause) As TStep

    Dim new_step As TStep
    
    If IsMissing(parent_clause) Then Set parent_clause = Nothing
    Set new_step = New TStep
    new_step.Parent = parent_clause
    new_step.parse_step_definition step_definition
    Set create_step = new_step
End Function
