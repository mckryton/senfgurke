VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_show_step_template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Then_the_resulting_report_will_show_the_code_templ_6F55B6204642(example_context As TContext)
    'Then the resulting report will show the code template only once
    Dim session As TSession
    Dim code_template_buffer As Collection

    Set session = example_context.get_value("session")
    Set code_template_buffer = session.Reporter.CodeTemplateBuffer
    TSpec.expect(code_template_buffer.Count).to_be 1
End Sub
