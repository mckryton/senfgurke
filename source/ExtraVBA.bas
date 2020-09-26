Attribute VB_Name = "ExtraVBA"
Option Explicit

'unfortunately there seems to be no compiler constant to distinct office applications
#Const APP_NAME = "Microsoft Excel"

Public Function existsItem(pvarKey As Variant, pcolACollection As Collection) As Boolean
                     
    On Error GoTo NOT_FOUND
    'use typename to access the collections item independ of its type (object or basic type)
    TypeName pcolACollection.Item(pvarKey)
    On Error GoTo 0
    existsItem = True
    Exit Function
                     
NOT_FOUND:
    existsItem = False
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

Function hash12(s As String)
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

Function hash4(txt)
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
