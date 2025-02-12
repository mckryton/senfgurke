Attribute VB_Name = "Excel_specific"
Option Explicit

Public Function get_this_document_path() As String
    get_this_document_path = ThisWorkbook.Path
End Function

