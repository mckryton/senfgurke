Attribute VB_Name = "ExtraVBA"
Option Explicit

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

Public Function collection_has_key(search_key As Variant, search_target As Collection) As Boolean
                     
    On Error GoTo NOT_FOUND
    'use typename to access the collections item independ of its type (object or basic type)
    TypeName search_target.Item(search_key)
    On Error GoTo 0
    collection_has_key = True
    Exit Function
                     
NOT_FOUND:
    collection_has_key = False
End Function

Public Function collection_has_value(search_value As Variant, search_target As Collection) As Boolean
                     
    Dim member_value As Variant
    Dim search_value_type As String
                     
    collection_has_value = False
    search_value_type = TypeName(search_value)
    For Each member_value In search_target
        If search_value_type = TypeName(member_value) Then
            If IsArray(search_value) Then
                If arrays_are_equal(search_value, member_value) Then
                    collection_has_value = True
                    Exit Function
                End If
            Else
                If member_value = search_value Then
                    collection_has_value = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function arrays_are_equal(first_array As Variant, second_array As Variant) As Boolean

    Dim first_value As String
    Dim second_value As String

    On Error GoTo UNSUPPORTED_ARRAY_ERROR
    first_value = Join(first_array, "#")
    second_value = Join(second_array, "#")
    On Error GoTo 0
    arrays_are_equal = False
    If first_value = second_value Then
        arrays_are_equal = True
    End If
    Exit Function
    
UNSUPPORTED_ARRAY_ERROR:
    Debug.Print "ExtraVBA ERROR: can't compare arrays with non-primitive values"
    Err.raise ERR_ID_UNSUPPORTED_ARRAY_ERROR, "ExtraVBA.arrays_are_equal", "can't compare arrays with non-primitive values"
End Function

Public Function get_path_separator() As String
    ' word and excel return path separator via Application.PathSeparator
    '  but this property is missing in Powerpoint

    #If MAC_OFFICE_VERSION >= 15 Then
        'in Office 2016 MAC M$ switched to / as path separator
        get_path_separator = "/"
    #ElseIf Mac Then
        get_path_separator = ":"
    #Else
        get_path_separator = "\"
    #End If
End Function

Public Function hash12(s As String)
    ' source: https://stackoverflow.com/questions/14717526/vba-hash-string
    ' create a 12 character hash from string s
    
    Dim l As Integer
    Dim l3 As Integer
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    
    l = Len(s)
    l3 = Int(l / 3)
    s1 = Mid(s, 1, l3)      ' first part
    s2 = Mid(s, l3 + 1, l3) ' middle part
    s3 = Mid(s, 2 * l3 + 1) ' the rest of the string...
    
    hash12 = hash4(s1) + hash4(s2) + hash4(s3)

End Function

Private Function hash4(txt As String)
    ' source: https://stackoverflow.com/questions/14717526/vba-hash-string
    Dim mask As Integer
    Dim j As Integer
    Dim nC As Integer
    Dim crc As Integer
    Dim c As String
    
    crc = &HFFFF
    
    For nC = 1 To Len(txt)
        j = Asc(Mid(txt, nC))
        crc = crc Xor j
        For j = 1 To 8
            mask = 0
            If crc / 2 <> Int(crc / 2) Then mask = &HA001
            crc = Int(crc / 2) And &H7FFF: crc = crc Xor mask
        Next j
    Next nC
    
    c = Hex$(crc)
    
    ' <<<<< new section: make sure returned string is always 4 characters long >>>>>
    ' pad to always have length 4:
    Do While Len(c) < 4
      c = "0" & c
    Loop
    
    hash4 = c
End Function

Public Function get_unix_timestamp(in_date As Date, in_time As Single) As Long

    Dim unix_ref_date As Double
    
    unix_ref_date = DateSerial(1970, 1, 1)
    get_unix_timestamp = CLng((in_date - unix_ref_date) * 86400 + (in_time * 1000))
End Function

Public Function get_unix_timestamp_now()
    get_unix_timestamp_now = get_unix_timestamp(Now, Timer)
End Function

Public Function trim_linebreaks(input_text As String) As String

    Dim result_text As String
    
    result_text = input_text
    Do While Left(result_text, 1) = vbLf
        result_text = Right(result_text, Len(result_text) - 1)
    Loop
    Do While Right(result_text, 1) = vbLf
        result_text = Left(result_text, Len(result_text) - 1)
    Loop
    trim_linebreaks = result_text
End Function
