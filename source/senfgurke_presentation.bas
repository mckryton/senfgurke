Attribute VB_Name = "senfgurke_presentation"
Option Explicit

Private Sub exportCode()

    Dim vbe_source_object As VBComponent
    Dim base_path As String
    Dim sub_path As String
    Dim file_path As String
    Dim path_separator As String
    Dim file_suffix As String

    path_separator = get_path_separator
    base_path = Application.Presentations("Senfgurke.pptm").Path
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
        If Left(vbe_source_object.name, 6) = "Steps_" Or Left(vbe_source_object.name, 8) = "Support_" Then
            sub_path = "step_definitions" & path_separator & "source    "
        Else
            sub_path = "source"
        End If
        file_path = base_path & path_separator & sub_path & path_separator & vbe_source_object.name & "." & file_suffix
        file_path = Replace(file_path, path_separator & "addins", vbNullString)
        Debug.Print "export code to " & file_path
        #If Mac Then
            'try not to change forms unless Microsoft offers full support for forms on the Mac!
            If Not file_suffix = "frm" Then
                vbe_source_object.Export file_path
            End If
        #Else
            vbe_source_object.Export file_path
        #End If
    Next
End Sub

