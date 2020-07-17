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
    #If MAC_OFFICE_VERSION >= 15 Then
        'in Office 2016 MAC M$ switched to / as path separator
        path_separator = "/"
    #ElseIf Mac Then
        path_separator = ":"
    #Else
        path_separator = "\"
    #End If
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
        If Left(vbe_source_object.Name, 1) = "T" Or vbe_source_object.Name = "ExtraVBA" Or vbe_source_object.Name = "Logger" Then
            sub_path = "source"
        Else
            sub_path = "test"
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
        export_logger.log "export code to " & file_path
    Next
    Exit Sub

error_handler:
    export_logger.log_function_error "ExtraVBA.exportCode"
End Sub
