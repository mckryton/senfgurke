Attribute VB_Name = "build_base"
Option Explicit

Global Const DOC_TYPE_ADDIN = "doc_type_addin"
Global Const DOC_TYPE_STEP_DEF = "doc_type_step_definition"

Public Sub build_addin()
    build_document DOC_TYPE_ADDIN, "Senfgurke", "source"
End Sub

Public Sub build_step_definitions()
    build_document DOC_TYPE_STEP_DEF, "Senfgurke_steps", "step_definitions"
End Sub

Private Sub build_document(doc_type As String, vba_project_name As String, source_dir_name As String)
    Const CODE_FILE_EXTENSIONS = "|.bas|.cls|.frm|"
    Dim addin_doc As Variant
    Dim source_path As String
    Dim file_name As String
    Dim file_extension As String
    Dim path_separator As String
    
    source_path = get_source_path(source_dir_name)
    If source_path = vbNullString Then
        Debug.Print "can't find source files - build FAILED"
        Exit Sub
    End If
    
    Set addin_doc = create_new_document(doc_type, vba_project_name)
    
    path_separator = get_path_separator
    file_name = Dir(source_path)
    While file_name <> vbNullString
        file_extension = Right(file_name, 4)
        If InStr(CODE_FILE_EXTENSIONS, file_extension) > 0 And Not is_exclusive_for_other_office_apps(file_name) Then
            addin_doc.VBProject.VBComponents.Import source_path & path_separator & file_name
            Debug.Print "imported>" & vbTab & file_name
        Else
            Debug.Print "ignored>" & vbTab & file_name
        End If
        file_name = Dir()
    Wend

    Debug.Print "build COMPLETED"
End Sub

Private Function get_source_path(source_path_sub_dir_name As String)
    Dim path_separator As String
    Dim build_dir As String
    Dim build_path As String
    Dim source_path As String
    Dim root_path As String
    
    'if the path to the Senfgurke repository is /home/user/senfgurke then the
    ' location of this document is expected to be /home/user/senfgurke/build while
    ' the source code for the AddIn is expected to be /home/user/senfgurke/source and
    ' the sorec code of the step definitions is expected to be /home/user/senfgurke/step_definitions
    path_separator = get_path_separator
    build_dir = path_separator & "build"
    'get the full path to the build path from this document
    build_path = get_build_path()
    If Right(build_path, Len(build_dir)) <> build_dir Then
        Debug.Print "ERROR: This document is expected to be located in the build sub dir but was found in >" & _
                get_build_path() & "<"
        get_source_path = vbNullString
    End If
    root_path = Left(build_path, Len(build_path) - Len(build_dir))
    get_source_path = root_path & path_separator & source_path_sub_dir_name
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

