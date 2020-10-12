Attribute VB_Name = "TConst"
Option Explicit

Public Const CLAUSE_TYPE_FEATURE = "feature"
Public Const CLAUSE_TYPE_RULE = "rule"
Public Const CLAUSE_TYPE_EXAMPLE = "example"

Public Const CLAUSE_ATTR_TYPE = "type"
Public Const CLAUSE_ATTR_NAME = "name"

Public Const ERR_ID_SCENARIO_SYNTAX_ERROR = vbError + 6010
Public Const ERR_ID_STEP_IS_STATUS_PENDING = vbError + 6020
Public Const ERR_ID_STEP_IS_STATUS_MISSING = vbError + 6030
Public Const ERR_ID_EXPECTATION_STATUS_FAILED = vbObjectError + 6500
Public Const ERR_ID_UNKNOWN_MSG_TYPE = vbObjectError + 6600

Public Const STATUS_OK = "OK"
Public Const STATUS_FAIL = "FAIL"
Public Const STATUS_MISSING = "MISSING"
Public Const STATUS_PENDING = "PENDING"

Public Const STEP_TYPE_GIVEN = "Given"
Public Const STEP_TYPE_WHEN = "When"
Public Const STEP_TYPE_THEN = "Then"

Public Const STEP_ATTR_TYPE = "step_type"
Public Const STEP_ATTR_HEAD = "step_head"
Public Const STEP_ATTR_NAME = "step_name"
Public Const STEP_ATTR_ERROR = "error"
