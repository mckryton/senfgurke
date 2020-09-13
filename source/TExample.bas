Attribute VB_Name = "TExample"
Option Explicit

Public Const STEP_TYPE_GIVEN = "Given"
Public Const STEP_TYPE_WHEN = "When"
Public Const STEP_TYPE_THEN = "Then"

Public Const STEP_ATTR_TYPE = "step_type"
Public Const STEP_ATTR_HEAD = "step_head"
Public Const STEP_ATTR_NAME = "step_name"
Public Const STEP_ATTR_ERROR = "error"

Public Const EXAMPLE_STEPS = "steps"

Public Function execute_step(step_attributes As Variant, feature_object As Variant) As Variant

    Dim step_result As Variant

    'TODO: refactor step implementation -> call steps indepent from theire class or module
    step_result = feature_object.run_step(step_attributes)
    execute_step = step_result
End Function
