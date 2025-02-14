Attribute VB_Name = "Excel_specific"
Option Explicit

Public Function get_Senfgurke_root_path() As String
    Dim root_path As String
    
    'assuming that the Senfgurke step definitions are located in the root dir of the Senfgurke repository
    get_Senfgurke_root_path = ThisWorkbook.Path
End Function

