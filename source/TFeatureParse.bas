Attribute VB_Name = "TFeatureParse"
Option Explicit

Public Function parse_feature(feature_text As String) As Collection

    Dim parsed_feature As Collection
    Dim lines As Variant
    Dim line As Variant
    Dim keyword_value As Variant
    Dim current_clause As Collection
    
    Set parsed_feature = New Collection
    lines = Split(feature_text, vbLf)
    If Not read_keyword_value(CStr(lines(0)))(0) = TFeature.CLAUSE_TYPE_FEATURE Then
        parsed_feature.Add "Feature lacks feature keyword at the beginning", "error"
    End If
    For Each line In lines
        keyword_value = read_keyword_value(CStr(line))
        If Not keyword_value(0) = vbNullString Then
            Set current_clause = New Collection
            current_clause.Add keyword_value(0), "type"
            current_clause.Add keyword_value(1), "name"
            parsed_feature.Add current_clause
        Else
            'PENDING add content to the current clause
        End If
    Next
    Set parse_feature = parsed_feature
End Function

Private Function read_keyword_value(text_line As String) As Variant

    Dim line_items As Variant
    Dim keyword As String
    
    line_items = Split(Trim(text_line), ":")
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
    read_keyword_value = Array(keyword, Trim(Right(text_line, Len(text_line) - InStr(text_line, ":"))))
End Function
