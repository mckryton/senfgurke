Attribute VB_Name = "TConfig"
Option Explicit

Public Property Get MaxStepFunctionNameLength() As Variant
    
    'VBA compiler crashes under MacOS (v16.41) if a function name is longer than 63 characters
    'VBA compiler crashes under MacOS (v16.75) if a function name is longer than 59 characters
    ' Note: if autocompile is active (default), it crashes automatically.
    ' specified max value is 255 characters
    MaxStepFunctionNameLength = 58
End Property
