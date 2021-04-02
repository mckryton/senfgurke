Attribute VB_Name = "TFeatureParser"
Option Explicit

Const KEYWORDS$ = "Example,Scenario,Scenario Outline,Rule,Ability,Business Needs,Feature,Background"

Public Function parse_feature(gherkin_text As String) As TFeature

    Dim parsed_feature As TFeature
    Dim lines As Variant
    Dim line As String
    Dim section_headline As Collection
    Dim current_section As Variant
    Dim line_index As Long
    Dim section_tags As Collection
    
    Set parsed_feature = parse_feature_definition(gherkin_text)
    If Not parsed_feature Is Nothing Then
        lines = Split(Trim(gherkin_text), vbLf)
        Set section_tags = clone_tags(parsed_feature.tags)
        Set current_section = Nothing
        For line_index = parsed_feature.ParsedLinesIndex + 1 To UBound(lines)
            line = Trim(lines(line_index))
            If is_section_definition_line(line) Then
                Set section_headline = read_section_headline(CStr(line))
                If Not current_section Is Nothing Then trim_description_linebreaks current_section
                Set current_section = create_section(parsed_feature, section_headline, section_tags, line)
                If Not section_headline("type") = SECTION_TYPE_BACKGROUND Then parsed_feature.sections.Add current_section
                If section_headline("type") = SECTION_TYPE_EXAMPLE Or section_headline("type") = SECTION_TYPE_BACKGROUND Then
                    line_index = parse_steps(lines, line_index, current_section)
                End If
                Set section_tags = clone_tags(parsed_feature.tags)
            ElseIf is_tag_line(line) Then
                add_tags line, section_tags
            ElseIf Not is_comment_line(line) Then
                'ignore comments
                If Not TypeName(current_section) = "TExample" Or TypeName(current_section) = "TBackground" Then
                    add_description CStr(line), parsed_feature
                End If
            End If
        Next
    End If
    Set parse_feature = parsed_feature
End Function

Private Sub trim_description_linebreaks(feature_section)

    Do While Right(feature_section.Description, 1) = vbLf
        feature_section.Description = Left(feature_section.Description, Len(feature_section.Description) - 1)
    Loop
    Do While Left(feature_section.Description, 1) = vbLf
        feature_section.Description = Right(feature_section.Description, Len(feature_section.Description) - 1)
    Loop
End Sub

Public Function parse_steps(feature_lines As Variant, example_start_index As Long, current_section As Variant) As Long

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
            Set current_step = current_section.Steps(current_section.Steps.Count)
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
        ElseIf is_step_line(line) Then
            current_section.Steps.Add create_step(line, current_section)
        ElseIf is_section_definition_line(line) Or Trim(line) = vbNullString Or is_tag_line(line) Then
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
    Dim feature_spec As Collection
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
        ElseIf is_comment_line(line) Or Trim(line) = vbNullString Then
            line_index = line_index + 1
        Else
            Set feature_spec = read_section_headline(line)
            If Not CStr(feature_spec("type")) = SECTION_TYPE_FEATURE Then
                Err.Raise ERR_ID_FEATURE_SYNTAX_ERROR, _
                            "TFeatureParser.parse_feature_definition", _
                            "Feature lacks feature keyword at the beginning"
            Else
                parsed_feature.Name = CStr(feature_spec("name"))
            End If
            header_finished = True
        End If
    Loop
    parsed_feature.ParsedLinesIndex = line_index
    Set parse_feature_definition = parsed_feature
End Function

Private Function is_section_definition_line(feature_line As String) As Boolean
    
    Dim keyword  As Variant
    Dim section_name As String
    
    'section definition format is /^<keyword>:[^:]*$/
    If UBound(Split(Trim(feature_line), ":")) > 0 Then
        section_name = Split(Trim(feature_line), ":")(0)
        For Each keyword In Split(KEYWORDS, ",")
            If section_name = CStr(keyword) Then
               is_section_definition_line = True
               Exit Function
            End If
        Next
    End If
    is_section_definition_line = False
End Function

Private Function is_comment_line(feature_line As String) As Boolean
    is_comment_line = Left(Trim(feature_line), 1) = "#"
End Function

Private Function is_tag_line(feature_line As String) As Boolean
    is_tag_line = Left(Trim(feature_line), 1) = "@"
End Function

Private Function is_step_line(feature_line As String) As Boolean

    Dim first_word As String
    
    first_word = Split(Trim(feature_line) & " ", " ")(0)
    is_step_line = InStr("Given When Then And But", first_word) > 0 And Len(first_word) > 2
End Function

Public Sub add_tags(feature_line As String, tags As Collection)

    Dim tag_list As Variant
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

Private Function create_section(parsed_feature As TFeature, section_headline As Collection, section_tags As Collection, feature_line As String) As Variant

    Dim current_section As Variant
    
    Select Case section_headline("type")
        Case SECTION_TYPE_RULE
            Set current_section = create_rule(rule_name:=section_headline("name"))
        Case SECTION_TYPE_EXAMPLE
            Set current_section = create_example(example_head:=section_headline("type"), example_name:=section_headline("name"), example_tags:=section_tags, feature_line:=feature_line)
        Case SECTION_TYPE_BACKGROUND
            'background already exists as an attribute of the TFeature class
            Set current_section = parsed_feature.Background
        Case Else
            Debug.Print "PARSE ERROR: unknown section >" & section_headline("type") & "<"
    End Select
    Set create_section = current_section
End Function

Private Function read_section_headline(text_line As String) As Collection

    Dim line_items As Variant
    Dim section_headline As Collection
    
    line_items = Split(Trim(text_line), ":")
    Set section_headline = New Collection
    If UBound(line_items) > 0 Then
        Select Case line_items(0)
            Case "Feature", "Ability", "Business Need"
                section_headline.Add SECTION_TYPE_FEATURE, "type"
            
            Case "Background"
                section_headline.Add SECTION_TYPE_BACKGROUND, "type"
            
            Case "Rule"
                section_headline.Add SECTION_TYPE_RULE, "type"
            
            Case "Scenario", "Scenario Outline", "Example"
                section_headline.Add SECTION_TYPE_EXAMPLE, "type"
            
            Case Else
                Debug.Print "PARSE ERROR: unknown keyword >" & line_items(0) & "<"
        End Select
    Else
        section_headline.Add "undefined", "type"
    End If
    section_headline.Add Trim(Right(text_line, Len(text_line) - InStr(text_line, ":"))), "name"
    Set read_section_headline = section_headline
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

    Dim section As Variant
    
    If feature.sections.Count = 0 Then
        If feature.Description = vbNullString Then
            feature.Description = Trim(line)
        Else
            feature.Description = feature.Description & vbLf & Trim(line)
        End If
    Else
        Set section = feature.sections(feature.sections.Count)
        If section.Description = vbNullString Then
            section.Description = Trim(line)
        Else
            section.Description = section.Description & vbLf & Trim(line)
        End If
    End If
End Sub

Private Function clone_tags(source_tags As Collection) As Collection

    Dim tag As Variant
    Dim target_tags As Collection
    
    Set target_tags = New Collection
    For Each tag In source_tags
        target_tags.Add tag, tag
    Next
    Set clone_tags = target_tags
End Function

Public Function create_step(step_definition As String, Optional parent_section) As TStep

    Dim new_step As TStep
    
    If IsMissing(parent_section) Then Set parent_section = New TExample
    Set new_step = New TStep
    new_step.Parent = parent_section
    new_step.parse_step_definition step_definition
    Set create_step = new_step
End Function
