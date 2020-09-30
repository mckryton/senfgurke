Attribute VB_Name = "TConst"
Option Explicit

Public Const CLAUSE_TYPE_FEATURE = "feature"
Public Const CLAUSE_TYPE_RULE = "rule"
Public Const CLAUSE_TYPE_EXAMPLE = "example"

Public Const CLAUSE_ATTR_TYPE = "type"
Public Const CLAUSE_ATTR_NAME = "name"

' custom error definitions
Public Const ERR_ID_SCENARIO_SYNTAX_ERROR = vbError + 6010
Public Const ERR_ID_STEP_IS_PENDING = vbError + 6020
Public Const ERR_ID_STEP_IS_MISSING = vbError + 6030
Public Const ERR_ID_EXPECTATION_FAILED = vbObjectError + 6500
Public Const ERR_ID_UNKNOWN_MSG_TYPE = vbObjectError + 6600

