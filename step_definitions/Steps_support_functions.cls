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

Dim m_date As Date
Dim m_seconds As Single
Dim m_unix_timestamp As Long

Public Sub Given_the_current_date_is_DBL_000001_18B10064814F(step_expressions As Collection)
    'Given the current date is {double} 00:00:01
    m_date = DateSerial(1970, 1, 1)
    m_seconds = 1
End Sub

Public Sub When_the_date_is_converted_to_an_unix_timestamp_3D49BC76BF30()
    'When the date is converted to an unix timestamp
    m_unix_timestamp = ExtraVBA.get_unix_timestamp(m_date, m_seconds)
End Sub

Public Sub Then_the_value_of_the_unix_timestamp_is_INT_DD8CBC7BB403(step_expressions As Collection)
    'Then the value of the unix timestamp is {integer}
    TSpec.expect(m_unix_timestamp).to_be step_expressions(1)
End Sub