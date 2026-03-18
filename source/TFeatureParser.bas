Attribute VB_Name = "TFeatureParser"
Option Explicit

Public Function parse_feature(gherkin_text As String) As TFeature

    Dim new_feature As TFeature
    Dim line As String
    Dim feature_lines As Variant
    Dim section_tags As Collection
    Dim new_rule As TRule
    Dim new_example As TExample
    Dim new_background As TBackground
    
    'On Error GoTo parsing_failed
    Set section_tags = New Collection
    Set new_feature = parse_feature_definition(gherkin_text)
    If new_feature Is Nothing Then
        TError.raise ERR_ID_FEATURE_SYNTAX_MISSING_DEFINITION, _
            "TFeatureParser.parse_feature"
    End If
    feature_lines = Split(gherkin_text, vbLf)
    Do While new_feature.parsed_lines < UBound(feature_lines)
        new_feature.parsed_lines = new_feature.parsed_lines + 1
        line = CStr(feature_lines(new_feature.parsed_lines))
        Select Case get_line_type(line)
            Case LINE_TYPE_COMMENT
                'ignore comments
            Case LINE_TYPE_TAGS
                parse_and_add_tags line, section_tags
            Case LINE_TYPE_BACKGROUND_START
                Set new_background = parse_background(gherkin_text, new_feature)
                Set new_feature.background = new_background
                'ignore any tags set before the background section
                Set section_tags = New Collection
            Case LINE_TYPE_RULE_START
                Set new_rule = parse_rule(gherkin_text, new_feature, section_tags)
                'ignore rules without examples
                If new_rule.sections.Count > 0 Then
                    new_rule.description = trim_description_linebreaks(new_rule.description)
                    new_feature.sections.Add new_rule
                    add_tags new_feature.tags, new_rule.tags
                End If
                Set section_tags = New Collection
            Case LINE_TYPE_EXAMPLE_START
                Set new_example = parse_example(gherkin_text, new_feature, section_tags)
                'ignore example without steps
                If new_example.steps.Count > 0 Then
                    new_example.description = trim_description_linebreaks(new_example.description)
                    new_feature.sections.Add new_example
                    add_tags new_feature.tags, new_example.tags
                End If
                Set section_tags = New Collection
            Case LINE_TYPE_DESCRIPTION
                new_feature.description = add_description(new_feature.description, line)
            Case Else
                'ignore empty lines after the first section (background, rule, or example)
                If Not Trim(line) = "" Then
                    TError.raise ERR_ID_FEATURE_SYNTAX_UNEXPECTED_LINE, _
                        "TFeatureParser.parse_feature", _
                        Array(new_feature.parsed_lines, line)
                End If
        End Select
    Loop
    new_feature.description = trim_description_linebreaks(new_feature.description)
    Set parse_feature = new_feature
    Exit Function
    
parsing_failed:
    'TODO handle exception one level up to get access to the feature file name
    Debug.Print "DEBUG parsing failed: " & Err.description
    Err.Clear
End Function

Public Function parse_background(gherkin_text As String, parent_feature As TFeature) As TBackground

    Dim feature_lines As Variant
    Dim line As String
    Dim new_background As TBackground

    feature_lines = Split(gherkin_text, vbLf)
    Set new_background = New TBackground
    Do While parent_feature.parsed_lines < UBound(feature_lines)
        parent_feature.parsed_lines = parent_feature.parsed_lines + 1
        line = CStr(feature_lines(parent_feature.parsed_lines))
        Select Case get_line_type(line)
            Case LINE_TYPE_COMMENT
                'ignore comments
            Case LINE_TYPE_STEP
                'parent_feature.parsed_lines = parse_steps(gherkin_text, parent_feature.parsed_lines, new_background.steps)
                parent_feature.parsed_lines = parent_feature.parsed_lines - 1
                Set new_background.steps = parse_step_list(gherkin_text, parent_feature)
                'steps will complete the background
                Set parse_background = new_background
                Exit Function
            Case LINE_TYPE_EXAMPLE_START, LINE_TYPE_RULE_START, LINE_TYPE_TAGS
                'new rules, examples or tags will also complete the background
                Set parse_background = new_background
                parent_feature.parsed_lines = parent_feature.parsed_lines - 1
                Exit Function
            Case LINE_TYPE_DESCRIPTION
                If new_background.steps.Count = 0 Then
                    new_background.description = add_description(new_background.description, line)
                Else
                    'ignore empty lines after the steps
                    If Not Trim(line) = "" Then
                        TError.raise ERR_ID_SECTION_SYNTAX_DESCRIPTION_AFTER_STEPS, _
                            "TFeatureParser.parse_background", _
                            Array(parent_feature.parsed_lines, line)
                    End If
                End If
            Case Else
                TError.raise ERR_ID_FEATURE_SYNTAX_UNEXPECTED_LINE, _
                    "TFeatureParser.parse_feature", _
                    Array(parent_feature.parsed_lines, line)
        End Select
    Loop
    Set parse_background = new_background
End Function

Public Function parse_rule(gherkin_text As String, parent_feature As TFeature, rule_tags As Collection) As TRule

    Dim feature_lines As Variant
    Dim line As String
    Dim new_example As TExample
    Dim new_rule As TRule
    Dim example_tags As Collection
    Dim parsed_tag_lines As Long

    Set example_tags = New Collection
    parsed_tag_lines = 0
    feature_lines = Split(gherkin_text, vbLf)
    line = CStr(feature_lines(parent_feature.parsed_lines))
    Set new_rule = create_rule(CStr(feature_lines(parent_feature.parsed_lines)), rule_tags, parent_feature)
    Do While parent_feature.parsed_lines < UBound(feature_lines)
        parent_feature.parsed_lines = parent_feature.parsed_lines + 1
        line = CStr(feature_lines(parent_feature.parsed_lines))
        Select Case get_line_type(line)
            Case LINE_TYPE_COMMENT
                'ignore comments
            Case LINE_TYPE_STEP
                'rules are not expected to have steps, they can have only examples
                TError.raise ERR_ID_SECTION_SYNTAX_STEPS_IN_RULE, _
                    "TFeatureParser.parse_rule", _
                    Array(parent_feature.parsed_lines, line)
            Case LINE_TYPE_TAGS
                parse_and_add_tags line, example_tags
                parsed_tag_lines = parsed_tag_lines + 1
            Case LINE_TYPE_EXAMPLE_START
                Set new_example = parse_example(gherkin_text, parent_feature, example_tags)
                Set example_tags = New Collection
                parsed_tag_lines = 0
                'ignore examples without steps
                If new_example.steps.Count > 0 Then
                    new_example.description = trim_description_linebreaks(new_example.description)
                    new_rule.sections.Add new_example
                    add_tags parent_feature.tags, new_example.tags
                    add_tags new_rule.tags, new_example.tags
                End If
            Case LINE_TYPE_RULE_START
                'another rule will finish this rule
                Set parse_rule = new_rule
                parent_feature.parsed_lines = parent_feature.parsed_lines - 1 - parsed_tag_lines
                Exit Function
            Case LINE_TYPE_DESCRIPTION
                If new_rule.sections.Count = 0 Then
                    new_rule.description = add_description(new_rule.description, line)
                Else
                    'ignore empty lines after the first example
                    If Not Trim(line) = "" Then
                        TError.raise ERR_ID_FEATURE_SYNTAX_UNEXPECTED_LINE, _
                            "TFeatureParser.parse_rule", _
                            Array(parent_feature.parsed_lines, line)
                    End If
                End If
        End Select
    Loop
    Set parse_rule = new_rule
End Function

Public Function parse_example(gherkin_text As String, parent_feature As TFeature, example_tags As Collection) As TExample
    'example is a synonym for scenario

    Dim feature_lines As Variant
    Dim line As String
    Dim new_example As TExample
    Dim new_outline As Collection
    Dim outline_table As TDataTable
    Dim outline_headline As Collection

    feature_lines = Split(gherkin_text, vbLf)
    line = CStr(feature_lines(parent_feature.parsed_lines))
    Set new_example = create_example(line, example_tags)
    Do While parent_feature.parsed_lines < UBound(feature_lines)
        parent_feature.parsed_lines = parent_feature.parsed_lines + 1
        line = CStr(feature_lines(parent_feature.parsed_lines))
        Select Case get_line_type(line)
            Case LINE_TYPE_COMMENT
                'ignore comments
            Case LINE_TYPE_STEP
                parent_feature.parsed_lines = parent_feature.parsed_lines - 1
                Set new_example.steps = TStepParser.parse_step_list(gherkin_text, parent_feature)
            Case LINE_TYPE_EXAMPLE_START, LINE_TYPE_RULE_START, LINE_TYPE_TAGS
                ' a new section or tags indicating the beginning of a new section will terminate the example definition
                parent_feature.parsed_lines = parent_feature.parsed_lines - 1
                Set parse_example = new_example
                Exit Function
            Case LINE_TYPE_OUTLINE_START
                Set new_outline = New Collection
                Set outline_headline = read_section_headline(line)
                new_outline.Add outline_headline.item("name"), "name"
                new_outline.Add Trim(line), "full_name"
                Set outline_table = TStepParser.parse_table(gherkin_text, parent_feature)
                new_outline.Add outline_table, "table"
                new_example.Outlines.Add new_outline
                'don't exit function here, because examples can have multiple outline tables
            Case LINE_TYPE_DESCRIPTION
                If new_example.steps.Count = 0 Then
                    new_example.description = add_description(new_example.description, line)
                Else
                    'ignore empty lines after the  example
                    If Not Trim(line) = "" Then
                        TError.raise ERR_ID_SECTION_SYNTAX_DESCRIPTION_AFTER_STEPS, _
                            "TFeatureParser.parse_example", _
                            Array(parent_feature.parsed_lines, line)
                    End If
                End If
            Case LINE_TYPE_TABLE_ROW
                TError.raise ERR_ID_STEP_SYNTAX_TABLE_WITHOUT_STEP, _
                    "TStepParser.parse_example"
            Case Else
                TError.raise ERR_ID_FEATURE_SYNTAX_UNEXPECTED_LINE, _
                    "TFeatureParser.parse_example", _
                    Array(parent_feature.parsed_lines, line)
        End Select
    Loop
    Set parse_example = new_example
End Function

Private Function trim_description_linebreaks(section_description As String) As String
    
    Dim trimmed_description As String
    
    trimmed_description = section_description
    Do While Right(trimmed_description, 1) = vbLf
        trimmed_description = Left(trimmed_description, Len(trimmed_description) - 1)
    Loop
    Do While Left(trimmed_description, 1) = vbLf
        trimmed_description = Right(trimmed_description, Len(trimmed_description) - 1)
    Loop
    trim_description_linebreaks = trimmed_description
End Function

Public Function parse_feature_definition(gherkin_text As String) As TFeature
    
    Dim parsed_feature As TFeature
    Dim line_index As Long
    Dim feature_lines As Variant
    Dim feature_spec As Collection
    Dim line As String

    Set parsed_feature = New TFeature
    feature_lines = Split(gherkin_text, vbLf)
    line_index = 0
    '>>> parse feature above feature name (tags and comments)
    Do While line_index <= UBound(feature_lines)
        line = CStr(feature_lines(line_index))
        Select Case get_line_type(line)
            Case LINE_TYPE_COMMENT
                'ignore comments
            Case LINE_TYPE_TAGS
                 parse_and_add_tags line, parsed_feature.tags
            Case LINE_TYPE_FEATURE_START
                Exit Do
            Case Else
                'ignore empty lines
                If Not Trim(line) = "" Then
                    TError.raise ERR_ID_FEATURE_SYNTAX_MISSING_FEATURE_KEYWORD, _
                        "TFeatureParser.parse_feature_definition"
                End If
        End Select
        line_index = line_index + 1
    Loop
    '>>> parse feature name <keyword>:<name>
    If line_index <= UBound(feature_lines) Then
        Set feature_spec = read_section_headline(line)
        parsed_feature.head = CStr(feature_spec("head"))
        parsed_feature.name = CStr(feature_spec("name"))
    Else
        TError.raise ERR_ID_FEATURE_SYNTAX_MISSING_FEATURE_NAME, _
                "TFeatureParser.parse_feature_definition"
    End If
    line_index = line_index + 1
    '>>> parse description below feature name
    Do While line_index <= UBound(feature_lines)
        line = CStr(feature_lines(line_index))
        Select Case get_line_type(line)
            Case LINE_TYPE_DESCRIPTION
                parsed_feature.description = add_description(parsed_feature.description, line)
            Case LINE_TYPE_COMMENT
                'ignore comments
            Case Else
                Exit Do
        End Select
        line_index = line_index + 1
    Loop
    parsed_feature.parsed_lines = line_index - 1
    Set parse_feature_definition = parsed_feature
End Function

Public Function get_line_type(line As String)

    Dim section_definition As Collection
    
    If is_step_line(line) Then
        get_line_type = LINE_TYPE_STEP
    ElseIf is_table_row(line) Then
        get_line_type = LINE_TYPE_TABLE_ROW
    ElseIf Trim(line) = """""""" Then
        get_line_type = LINE_TYPE_DOCSTRING_LIMIT
    ElseIf is_comment_line(line) Then
        get_line_type = LINE_TYPE_COMMENT
    ElseIf is_tag_line(line) Then
        get_line_type = LINE_TYPE_TAGS
    ElseIf is_section_definition_line(line) Then
        Set section_definition = read_section_headline(line)
        Select Case section_definition("type")
            Case SECTION_TYPE_RULE
                get_line_type = LINE_TYPE_RULE_START
            Case SECTION_TYPE_EXAMPLE
                get_line_type = LINE_TYPE_EXAMPLE_START
            Case SECTION_TYPE_BACKGROUND
                get_line_type = LINE_TYPE_BACKGROUND_START
            Case SECTION_TYPE_FEATURE
                get_line_type = LINE_TYPE_FEATURE_START
            Case SECTION_TYPE_OUTLINE
                get_line_type = LINE_TYPE_OUTLINE_START
        End Select
    Else
        get_line_type = LINE_TYPE_DESCRIPTION
    End If
End Function

Private Function is_section_definition_line(feature_line As String) As Boolean
    
    Dim keyword  As Variant
    Dim section_name As String
    
    'section definition format is /^<keyword>:[^:]*$/
    If UBound(Split(Trim(feature_line), ":")) > 0 Then
        section_name = Split(Trim(feature_line), ":")(0)
        For Each keyword In Split(START_SECTION_KEYWORDS, ",")
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
    is_step_line = InStr(" Given When Then And But ", " " & first_word & " ") > 0 And Len(first_word) > 2
End Function

Private Function is_table_row(feature_line As String) As Boolean
    is_table_row = False
    If Left(Trim(feature_line), 1) = "|" Then
        If Right(Trim(feature_line), 1) <> "|" Then
            TError.raise ERR_ID_STEP_SYNTAX_INCOMPLETE_TABLE_ROW, _
                "TFeatureParser.is_table_row", _
                Array(Trim(feature_line))
        End If
        is_table_row = True
    End If
End Function

Public Sub parse_and_add_tags(feature_line As String, tags As Collection)

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

Private Sub add_tags(source_tags As Collection, target_tags As Collection)

    Dim tag As Variant
    
    For Each tag In source_tags
        If Not ExtraVBA.collection_has_value(tag, target_tags) Then
            target_tags.Add tag, tag
        End If
    Next
End Sub

Private Function read_section_headline(text_line As String) As Collection

    Dim line_items As Variant
    Dim section_headline As Collection
    
    line_items = Split(Trim(text_line), ":")
    Set section_headline = New Collection
    If UBound(line_items) > 0 Then
        section_headline.Add Trim(line_items(0)), "head"
        Select Case line_items(0)
            Case "Feature", "Ability", "Business Need"
                section_headline.Add SECTION_TYPE_FEATURE, "type"
            Case "Background"
                section_headline.Add SECTION_TYPE_BACKGROUND, "type"
            Case "Rule"
                section_headline.Add SECTION_TYPE_RULE, "type"
            Case "Scenario", "Scenario Outline", "Example"
                section_headline.Add SECTION_TYPE_EXAMPLE, "type"
            Case "Examples"
                section_headline.Add SECTION_TYPE_OUTLINE, "type"
            Case Else
                Debug.Print "PARSE ERROR: unknown keyword >" & line_items(0) & "<"
        End Select
    Else
        section_headline.Add "undefined", "type"
    End If
    section_headline.Add Trim(Right(text_line, Len(text_line) - InStr(text_line, ":"))), "name"
    Set read_section_headline = section_headline
End Function

Private Function create_rule(line As String, rule_tags As Collection, feature_parent As TFeature) As TRule

    Dim new_rule As TRule
    Dim section_definition As Collection
    
    Set new_rule = New TRule
    Set new_rule.tags = rule_tags
    Set new_rule.parent = feature_parent
    Set section_definition = read_section_headline(line)
    new_rule.name = section_definition("head")
    new_rule.name = section_definition("name")
    Set create_rule = new_rule
End Function

Private Function create_example(line As String, example_tags As Collection) As TExample

    Dim new_example As TExample
    Dim section_definition As Collection
    
    Set new_example = New TExample
    new_example.tags = example_tags
    Set section_definition = read_section_headline(line)
    new_example.head = section_definition("head")
    new_example.name = section_definition("name")
    new_example.OriginalHeadline = Trim(line)
    Set create_example = new_example
End Function

Private Function add_description(section_description As String, line As String) As String

    'ignore empty lines at description start
    If section_description <> vbNullString Or Trim(line) <> vbNullString Then
        'add linebreak starting with the second line
        If section_description <> vbNullString Then
            add_description = section_description & vbLf & Trim(line)
        Else
            add_description = Trim(line)
        End If
    Else
        add_description = vbNullString
    End If
End Function

