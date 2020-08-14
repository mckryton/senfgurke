Attribute VB_Name = "TFeatureImport"

Public Sub read_text()

    Dim strFilename As String
    Dim strTextLine As String
    Dim iFile As Integer
    
    iFile = FreeFile
    strFilename = "/Users/mac/source/vba/senfgurke/features/execute_example_steps.feature"
    Open strFilename For Input As #iFile
    Do Until EOF(1)
        Line Input #1, strTextLine
        Debug.Print strTextLine
    Loop
    Close #iFile
End Sub
