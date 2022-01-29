VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_support_functions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_the_current_date_is_01011970_000001_515850950AD6(example_context As TContext)
    'Given the current date is 01.01.1970 00:00:01
    example_context.set_value DateSerial(1970, 1, 1), "sample_date"
    example_context.set_value 1, "sample_seconds"
End Sub

Public Sub When_the_date_is_converted_to_an_unix_timestamp_3D49BC76BF30(example_context As TContext)
    'When the date is converted to an unix timestamp
    Dim unix_timestamp As Long
    Dim sample_date As Date
    Dim sample_seconds As Single
    
    sample_date = example_context.get_value("sample_date")
    sample_seconds = example_context.get_value("sample_seconds")
    unix_timestamp = ExtraVBA.get_unix_timestamp(sample_date, sample_seconds)
    example_context.set_value unix_timestamp, "unix_timestamp"
End Sub

Public Sub Then_the_value_of_the_unix_timestamp_is_INT_DD8CBC7BB403(example_context As TContext, step_expressions As Collection)
    'Then the value of the unix timestamp is {integer}
    Dim unix_timestamp As Long
    
    unix_timestamp = example_context.get_value("unix_timestamp")
    TSpec.expect(unix_timestamp).to_be step_expressions(1)
End Sub
