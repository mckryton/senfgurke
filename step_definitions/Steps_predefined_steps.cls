VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_predefined_steps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' To prevent autoomatically generated code this class provides some steps with
' a pre-defined behavior like failingor pending.

Option Explicit

Public Sub Given_a_valid_step_6A35DF3A18EC(example_context As TContext)
    'Given a valid step
End Sub

Public Sub Given_an_invalid_step_FF6BEF10050E(example_context As TContext)
    'And an invalid step
    TSpec.expect(1).to_be 2
End Sub

Public Sub Given_a_pending_step_6A3533BBBFE5(example_context As TContext)
    pending
End Sub

Public Sub Given_an_empty_step_6A35DCCE89C0(example_context As TContext)
    'Given an empty step
    
End Sub
