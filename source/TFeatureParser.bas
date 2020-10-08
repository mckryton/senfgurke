Attribute VB_Name = "TFeatureParser"
Option Explicit

Public Function parse_feature(feature_text As String) As TFeature

    Dim parsed_feature As TFeature
    Dim lines As Variant
    Dim line As String
    Dim keyword_value As Variant
    Dim current_clause As Variant
    Dim step_head_name As Variant
    Dim has_example_steps As Boolean
    Dim line_index As Long
    Dim clause_tags As Collection
    
    Set parsed_feature = parse_feature_definition(feature_text)
    Set current_clause = Nothing
    lines = Split(Trim(feature_text), vbLf)
    Set clause_tags = clone_tags(parsed_feature.Tags)
    If parsed_feature.ErrorStatus = vbNullString Then
        For line_index = parsed_feature.ParsedLinesIndex + 1 To UBound(lines)
            line = Trim(lines(line_index))
            keyword_value = read_keyword_value(CStr(line))
            If Not keyword_value(0) = vbNullString Then
                Set current_clause = update_feature(parsed_feature, CStr(keyword_value(0)), CStr(keyword_value(1)), clause_tags, line)
                has_example_steps = False
                Set clause_tags = clone_tags(parsed_feature.Tags)
            ElseIf is_tag_line(line) Then
                add_tags line, clause_tags
            Else
                If TypeName(current_clause) = "TExample" Then
                    step_head_name = read_step_head_name(CStr(line))
                    If step_head_name(0) <> vbNullString Then
                        current_clause.add_step step_head:=CStr(step_head_name(0)), step_name:=CStr(step_head_name(1))
                        has_example_steps = True
                    End If
                End If
                If Not has_example_steps Then
                    add_description CStr(line), parsed_feature
                End If
            End If
        Next
    End If
    Set parse_feature = parsed_feature
End Function

Private Function parse_feature_definition(feature_text As String) As TFeature
    
    Dim parsed_feature As TFeature
    Dim line_index As Long
    Dim feature_lines As Variant
    Dim header_finished As Boolean
    Dim feature_spec As Variant
    Dim line As String

    Set parsed_feature = New TFeature
    header_finished = False
    line_index = 0
    feature_lines = Split(feature_text, vbLf)
    Do While line_index < UBound(feature_lines) And header_finished = False
        line = CStr(feature_lines(line_index))
        If is_tag_line(line) Then
             parsed_feature.add_tags line
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

Private Function is_tag_line(feature_line As String) As Boolean

    If Left(Trim(feature_line), 1) = "@" Then
        is_tag_line = True
    Else
        is_tag_line = False
    End If
End Function

Public Sub add_tags(feature_line As String, Tags As Collection)

    Dim tag_list As Variant
    Dim index As Long
    Dim tag As Variant
    
    tag_list = Split(Trim(feature_line), " ")
    For Each tag In tag_list
        If Len(tag) > 1 Then
            If Left(tag, 1) = "@" Then
                If Not ExtraVBA.existsItem(tag, Tags) Then
                    Tags.Add tag, tag
                End If
            End If
        End If
    Next
End Sub

Private Function update_feature(parsed_feature As TFeature, clause_keyword As String, clause_name As String, clause_tags As Collection, feature_line As String)

    Dim current_clause As Variant
    
    Select Case clause_keyword
        Case CLAUSE_TYPE_RULE
            Set current_clause = create_rule(rule_name:=clause_name)
            parsed_feature.Clauses.Add current_clause
        Case CLAUSE_TYPE_EXAMPLE
            Set current_clause = create_example(example_head:=CStr(clause_keyword), example_name:=clause_name, example_tags:=clause_tags, feature_line:=feature_line)
            parsed_feature.Clauses.Add current_clause
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
            Case "Rule"
                keyword = CLAUSE_TYPE_RULE
            Case "Scenario"
                keyword = CLAUSE_TYPE_EXAMPLE
            Case "Scenario Outline"
                keyword = CLAUSE_TYPE_EXAMPLE
            Case "Example"
                keyword = CLAUSE_TYPE_EXAMPLE
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
    new_example.Tags = example_tags
    new_example.OriginalHeadline = feature_line
    Set create_example = new_example
End Function

Private Sub add_description(line As String, feature As TFeature)

    Dim clause As Variant
    
    If feature.Clauses.Count = 0 Then
        If feature.description = vbNullString Then
            feature.description = Trim(line)
        Else
            feature.description = feature.description & vbLf & Trim(line)
        End If
    Else
        Set clause = feature.Clauses(feature.Clauses.Count)
        If clause.description = vbNullString Then
            clause.description = Trim(line)
        Else
            clause.description = clause.description & vbLf & Trim(line)
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

Public Function parse_feature_list(features_as_text As Collection) As Collection
    
    Dim feature_text As Variant
    Dim parsed_features As Collection

    Set parsed_features = New Collection
    For Each feature_text In features_as_text
        parsed_features.Add parse_feature(CStr(feature_text))
    Next
    Set parse_feature_list = parsed_features
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
