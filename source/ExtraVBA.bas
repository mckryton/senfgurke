Attribute VB_Name = "ExtraVBA"
Option Explicit

'unfortunately there seems to be no compiler constant to distinct office applications
#Const APP_NAME = "Microsoft Excel"

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
    If first_value = second_value Then
        arrays_are_equal = True
    Else
        arrays_are_equal = False
    End If
    Exit Function
    
UNSUPPORTED_ARRAY_ERROR:
    Debug.Print "ExtraVBA ERROR: can't compare arrays with non-primitive values"
    Err.Raise ERR_ID_UNSUPPORTED_ARRAY_ERROR, "ExtraVBA.arrays_are_equal", "can't compare arrays with non-primitive values"
End Function

Private Sub exportCode()

    Dim vbe_source_object As VBComponent
    Dim base_path As String
    Dim sub_path As String
    Dim file_path As String
    Dim path_separator As String
    Dim file_suffix As String
    Dim export_logger As Logger

    On Error GoTo error_handler
    Set export_logger = New Logger
    path_separator = get_path_separator
    #If APP_NAME = "Microsoft Powerpoint" Then
        base_path = ActivePresentation.Path
    #ElseIf APP_NAME = "Microsoft Excel" Then
        base_path = ThisWorkbook.Path
    #End If
    For Each vbe_source_object In Application.VBE.VBProjects("Senfgurke").VBComponents
        Select Case vbe_source_object.Type
            Case vbext_ct_StdModule
                file_suffix = "bas"
            Case vbext_ct_ClassModule
                file_suffix = "cls"
            Case vbext_ct_Document
                file_suffix = "cls"
            Case vbext_ct_MSForm
                file_suffix = "frm"
            Case Else
                file_suffix = "txt"
        End Select
        If Left(vbe_source_object.Name, 6) = "Steps_" Or Left(vbe_source_object.Name, 8) = "Support_" Then
            sub_path = "step_definitions"
        Else
            sub_path = "source"
        End If
        file_path = base_path & path_separator & sub_path & path_separator & vbe_source_object.Name & "." & file_suffix
        #If Mac Then
            'try not to change forms unless Microsoft offers full support for forms on the Mac!
            If Not file_suffix = "frm" Then
                vbe_source_object.Export file_path
            End If
        #Else
            vbe_source_object.Export file_path
        #End If
        export_logger.Log "export code to " & file_path
    Next
    Exit Sub

error_handler:
    export_logger.log_function_error "ExtraVBA.exportCode"
End Sub

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
    
    Dim l As Integer, l3 As Integer
    Dim s1 As String, s2 As String, s3 As String
    
    l = Len(s)
    l3 = Int(l / 3)
    s1 = Mid(s, 1, l3)      ' first part
    s2 = Mid(s, l3 + 1, l3) ' middle part
    s3 = Mid(s, 2 * l3 + 1) ' the rest of the string...
    
    hash12 = hash4(s1) + hash4(s2) + hash4(s3)

End Function

Private Function hash4(txt)
    ' source: https://stackoverflow.com/questions/14717526/vba-hash-string
    Dim x As Long
    Dim mask, i, j, nC, crc As Integer
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
    While Len(c) < 4
      c = "0" & c
    Wend
    
    hash4 = c
End Function

Public Function get_unix_timestamp(in_date As Date, in_time As Single) As Long

    Dim unix_ref_date As Double
    
    unix_ref_date = DateSerial(1970, 1, 1)
    get_unix_timestamp = CLng((in_date - unix_ref_date) * 86400 + (in_time * 1000))
End Function
