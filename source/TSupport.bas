Attribute VB_Name = "TSupport"
'this module is for sharing helper functions between your test cases

Option Explicit

Public Function create_config_slide(pConfigPresentation As Presentation, Optional pRuleName) As Slide
    
    Dim validator_pres As Presentation
    Dim config_slide As Slide
    Dim rule_name As String
    Dim config_template_index As Integer
    Dim config_table As shape
    
    If IsMissing(pRuleName) Then
        rule_name = "Rule: RuleName"
    Else
        rule_name = "Rule: " & pRuleName
    End If
    Set validator_pres = Application.Presentations("SlideValidator.pptm")
    config_template_index = get_config_template_index(pConfigPresentation)
    Set config_slide = pConfigPresentation.Slides.AddSlide(1, pConfigPresentation.SlideMaster.CustomLayouts(config_template_index))
    config_slide.Shapes.Title.TextFrame.TextRange.Text = rule_name
    config_slide.Name = rule_name
    '.AddTable returns the shape that contains the table not the table itself
    Set config_table = config_slide.Shapes.AddTable(1, 3, ExtraPpt.cm2points(2.33), ExtraPpt.cm2points(4.7), ExtraPpt.cm2points(29.21), ExtraPpt.cm2points(2.42))
    config_table.Table.Cell(1, 1).shape.TextFrame.TextRange.Text = "Parameter"
    config_table.Table.Cell(1, 2).shape.TextFrame.TextRange.Text = "Value"
    config_table.Table.Cell(1, 3).shape.TextFrame.TextRange.Text = "Description"
    Set create_config_slide = config_slide
End Function

Public Function create_slide_validator_pres() As Presentation

    Dim config_presentation As Presentation
    Dim slide_validator As Presentation
    Dim master_slide As CustomLayout
    Dim config_template_index As Integer
        
    Set slide_validator = Application.Presentations("SlideValidator.pptm")
    Set config_presentation = Application.Presentations.Add
    
    config_template_index = get_config_template_index(slide_validator)
    Set master_slide = slide_validator.SlideMaster.CustomLayouts(config_template_index)
    master_slide.Copy
    config_presentation.SlideMaster.CustomLayouts.Paste
    
    Set create_slide_validator_pres = config_presentation
End Function

Public Function get_config_table(pConfigSlide As Slide) As Table
    
    Dim config_table As Table
    Dim slide_shape As shape
    
    Set get_config_table = Nothing
    If LCase(Left(Trim(pConfigSlide.Shapes.Title.TextFrame.TextRange.Text), 4)) <> "rule" Then
        Exit Function
    End If
    For Each slide_shape In pConfigSlide.Shapes
        If slide_shape.HasTable Then
            If slide_shape.Table.Columns.Count >= 3 Then
                If LCase(Trim(slide_shape.Table.Cell(1, 1).shape.TextFrame.TextRange.Text)) = "parameter" _
                  And LCase(Trim(slide_shape.Table.Cell(1, 2).shape.TextFrame.TextRange.Text)) = "value" _
                  And LCase(Trim(slide_shape.Table.Cell(1, 3).shape.TextFrame.TextRange.Text)) = "description" Then
                    Set get_config_table = slide_shape.Table
                End If
            End If
        End If
    Next
End Function

Public Sub add_config_parameter(pConfigTable As Table, pConfigParameter As Variant)

    Dim parameter_row As Row
    
    Set parameter_row = pConfigTable.Rows.Add
    parameter_row.Cells(1).shape.TextFrame.TextRange.Text = pConfigParameter(0)
    parameter_row.Cells(2).shape.TextFrame.TextRange.Text = pConfigParameter(1)
    parameter_row.Cells(3).shape.TextFrame.TextRange.Text = pConfigParameter(2)
End Sub

Private Function get_config_template_index(pConfigPresentation As Presentation) As Integer

    Dim custom_layout As CustomLayout
    
    For Each custom_layout In pConfigPresentation.SlideMaster.CustomLayouts
        If custom_layout.Name = Validator.CONFIG_TEMPLATE_NAME Then
            get_config_template_index = custom_layout.index
            Exit Function
        End If
    Next
    Err.raise Validator.ERR_ID_MISSING_CFG_MASTER_SLIDE, description:="couldn't find a custom layout named >" & _
                Validator.CONFIG_TEMPLATE_NAME & "< in presentation " & pConfigPresentation.Name
End Function

Public Function add_slide_with_textbox(target_presentation As Presentation, font_name As String, Optional shape_name) As Slide

    Dim slide_with_textbox As Slide
    Dim textbox As shape

    Set slide_with_textbox = add_empty_slide(target_presentation)
    Set textbox = slide_with_textbox.Shapes.AddTextbox(msoTextOrientationHorizontal, 200, 200, 400, 200)
    textbox.TextFrame.TextRange.font.Name = font_name
    textbox.TextFrame.TextRange.Text = "This text is using " & font_name & " as font."
    If Not IsMissing(shape_name) Then
        textbox.Name = shape_name
    End If
    Set add_slide_with_textbox = slide_with_textbox
End Function

Public Function add_empty_slide(target_presentation As Presentation) As Slide

    Set add_empty_slide = target_presentation.Slides.AddSlide(1, target_presentation.SlideMaster.CustomLayouts(7))
End Function

Public Sub close_test_presentations()

    Dim current_pres As Presentation
    
    For Each current_pres In Application.Presentations
        If current_pres.Name <> "SlideValidator.pptm" Then
            current_pres.Saved = msoTrue
            current_pres.Close
        End If
    Next
End Sub
