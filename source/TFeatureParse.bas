Attribute VB_Name = "TFeatureParse"
Option Explicit

Public Function parse_feature(feature_text As String) As Collection

    Dim parsed_feature As Collection
    Dim lines As Variant
    Dim line As Variant
    Dim keyword_value As Variant
    Dim current_clause As Collection
    Dim step_head_name As Variant
    Dim is_example_started As Boolean
    
    Set parsed_feature = New Collection
    lines = Split(feature_text, vbLf)
    If Not read_keyword_value(CStr(lines(0)))(0) = TFeature.CLAUSE_TYPE_FEATURE Then
        parsed_feature.Add "Feature lacks feature keyword at the beginning", "error"
    Else
        For Each line In lines
            keyword_value = read_keyword_value(CStr(line))
            If Not keyword_value(0) = vbNullString Then
                Set current_clause = New Collection
                current_clause.Add keyword_value(0), "type"
                current_clause.Add keyword_value(1), "name"
                parsed_feature.Add current_clause
                is_example_started = False
            Else
                If current_clause("type") = TFeature.CLAUSE_TYPE_EXAMPLE Then
                    step_head_name = read_step_head_name(CStr(line))
                    If step_head_name(0) <> vbNullString Then
                        add_step step_head_name, current_clause
                        is_example_started = True
                    End If
                End If
                If Not is_example_started Then
                    add_description CStr(line), current_clause
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
                keyword = TFeature.CLAUSE_TYPE_FEATURE
            Case "Ability"
                keyword = TFeature.CLAUSE_TYPE_FEATURE
            Case "Business Need"
                keyword = TFeature.CLAUSE_TYPE_FEATURE
            Case "Rule"
                keyword = TFeature.CLAUSE_TYPE_RULE
            Case "Scenario"
                keyword = TFeature.CLAUSE_TYPE_EXAMPLE
            Case "Scenario Outline"
                keyword = TFeature.CLAUSE_TYPE_EXAMPLE
            Case "Example"
                keyword = TFeature.CLAUSE_TYPE_EXAMPLE
        End Select
    End If
    read_keyword_value = Array(keyword, trim(Right(text_line, Len(text_line) - InStr(text_line, ":"))))
End Function

Private Sub add_description(line As String, clause As Collection)

    Dim description As String

    If Not ExtraVBA.existsItem("description", clause) Then
        clause.Add trim(line), "description"
    Else
        description = clause("description") & vbLf & trim(line)
        clause.Remove "description"
        clause.Add description, "description"
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

Private Sub add_step(step_head_name As Variant, clause As Collection)

    Dim step_head As String
    Dim step_name As String
    Dim steps As Collection
    Dim new_step As Collection
    Dim previous_step_type As String
    
    step_head = step_head_name(0)
    step_name = step_head_name(1)
    If Not ExtraVBA.existsItem("steps", clause) Then
        Set steps = New Collection
        clause.Add steps, "steps"
        previous_step_type = TExample.STEP_TYPE_GIVEN
    Else
        Set steps = clause("steps")
        previous_step_type = steps(steps.Count)("type")
    End If
    Set new_step = New Collection
    If step_head = "And" Or step_head = "But" Then
        new_step.Add previous_step_type, "type"
    Else
        new_step.Add get_step_type(step_head), "type"
    End If
    new_step.Add step_head, "head"
    new_step.Add step_name, "name"
    steps.Add new_step
End Sub

Private Function get_step_type(step_head As String) As String

    Select Case step_head
        Case "Given"
            get_step_type = TExample.STEP_TYPE_GIVEN
        Case "When"
            get_step_type = TExample.STEP_TYPE_WHEN
        Case "Then"
            get_step_type = TExample.STEP_TYPE_THEN
        Case Else
            get_step_type = "unknown step type"
    End Select
End Function
