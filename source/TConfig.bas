Attribute VB_Name = "TConfig"
Option Explicit

Public Property Get StepImplementations() As Variant
    
    'add all classes with step implmentations here:
    StepImplementations = Array(New Steps_Run_Examples, New Steps_Load_Feature_Files, New Steps_Run_Steps)
End Property

Public Property Get MaxStepFunctionNameLength() As Variant
    
    'VBA compiler crashes under MacOS (v16.41) if a function name is longer than 63 characters
    ' specified max value is 255 characters
    MaxStepFunctionNameLength = 63
End Property
