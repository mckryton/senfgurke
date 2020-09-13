Attribute VB_Name = "TFeatureParser"
Option Explicit

Public Function parse_feature(feature_text As String) As TFeature

    Dim parsed_feature As TFeature
    Dim lines As Variant
    Dim line As Variant
    Dim keyword_value As Variant
    Dim current_clause As Variant
    Dim step_head_name As Variant
    Dim is_example_started As Boolean
    
    Set parsed_feature = New TFeature
    Set current_clause = Nothing
    lines = Split(feature_text, vbLf)
    If Not read_keyword_value(CStr(lines(0)))(0) = TFeatureRunner.CLAUSE_TYPE_FEATURE Then
        parsed_feature.ErrorStatus = "Feature lacks feature keyword at the beginning"
    Else
        For Each line In lines
            keyword_value = read_keyword_value(CStr(line))
            If Not keyword_value(0) = vbNullString Then
                Select Case keyword_value(0)
                    Case TFeatureRunner.CLAUSE_TYPE_FEATURE
                        parsed_feature.Head = CStr(keyword_value(0))
                        parsed_feature.Name = CStr(keyword_value(1))
                    Case TFeatureRunner.CLAUSE_TYPE_RULE
                        Set current_clause = create_rule(rule_name:=CStr(keyword_value(1)))
                        parsed_feature.Clauses.Add current_clause
                    Case TFeatureRunner.CLAUSE_TYPE_EXAMPLE
                        Set current_clause = create_example(example_head:=CStr(keyword_value(0)), example_name:=CStr(keyword_value(1)))
                        parsed_feature.Clauses.Add current_clause
                End Select
                is_example_started = False
            Else
                If TypeName(current_clause) = "TExample" Then
                    step_head_name = read_step_head_name(CStr(line))
                    If step_head_name(0) <> vbNullString Then
                        current_clause.add_step step_head:=CStr(step_head_name(0)), step_name:=CStr(step_head_name(1))
                        is_example_started = True
                    End If
                End If
                If Not is_example_started Then
                    add_description CStr(line), parsed_feature
                End If
            End If
        Next
    End If
    Set parse_feature = parsed_feature
End Function

Private Function read_keyword_value(text_line As String) As Variant

    Dim line_items As Variant
    Dim keyword As String
    
    line_items = Split(trim(text_line), ":")
    keyword = vbNullString
    If UBound(line_items) > 0 Then
        Select Case line_items(0)
            Case "Feature"
                keyword = TFeatureRunner.CLAUSE_TYPE_FEATURE
            Case "Ability"
                keyword = TFeatureRunner.CLAUSE_TYPE_FEATURE
            Case "Business Need"
                keyword = TFeatureRunner.CLAUSE_TYPE_FEATURE
            Case "Rule"
                keyword = TFeatureRunner.CLAUSE_TYPE_RULE
            Case "Scenario"
                keyword = TFeatureRunner.CLAUSE_TYPE_EXAMPLE
            Case "Scenario Outline"
                keyword = TFeatureRunner.CLAUSE_TYPE_EXAMPLE
            Case "Example"
                keyword = TFeatureRunner.CLAUSE_TYPE_EXAMPLE
        End Select
    End If
    read_keyword_value = Array(keyword, trim(Right(text_line, Len(text_line) - InStr(text_line, ":"))))
End Function

Private Function create_rule(rule_name As String) As TRule

    Dim new_rule As TRule
    
    Set new_rule = New TRule
    new_rule.Name = rule_name
    Set create_rule = new_rule
End Function

Private Function create_example(example_head As String, example_name As String) As TExample

    Dim new_example As TExample
    
    Set new_example = New TExample
    new_example.Head = example_head
    new_example.Name = example_name
    Set create_example = new_example
End Function

Private Sub add_description(line As String, feature As TFeature)

    Dim clause As Variant
    
    If feature.Clauses.Count = 0 Then
        If feature.Description = vbNullString Then
            feature.Description = trim(line)
        Else
            feature.Description = feature.Description & vbLf & trim(line)
        End If
    Else
        Set clause = feature.Clauses(feature.Clauses.Count)
        If clause.Description = vbNullString Then
            clause.Description = trim(line)
        Else
            clause.Description = clause.Description & vbLf & trim(line)
        End If
    End If
End Sub

Private Function read_step_head_name(line As String) As Variant

    Dim step_type As String
    Dim trim_line As String

    trim_line = trim(line)
    If Len(trim_line) > 0 Then
        step_type = Split(trim_line, " ")(0)
        If InStr("Given When Then And But", step_type) > 0 Then
            read_step_head_name = Array(step_type, Right(trim_line, Len(trim_line) - InStr(trim_line, " ")))
            Exit Function
        End If
    End If
    read_step_head_name = Array(vbNullString, vbNullString)
End Function

