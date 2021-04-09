Attribute VB_Name = "TConst"
Option Explicit

Public Const SECTION_TYPE_FEATURE$ = "feature"
Public Const SECTION_TYPE_RULE$ = "rule"
Public Const SECTION_TYPE_EXAMPLE$ = "example"
Public Const SECTION_TYPE_BACKGROUND$ = "background"

Public Const SECTION_ATTR_TYPE$ = "type"
Public Const SECTION_ATTR_NAME$ = "name"

Public Const ERR_ID_FEATURE_SYNTAX_ERROR& = vbError + 6010
Public Const ERR_ID_SCENARIO_SYNTAX_ERROR& = vbError + 6012
Public Const ERR_ID_STEP_IS_STATUS_PENDING& = vbError + 6020
Public Const ERR_ID_STEP_IS_STATUS_MISSING& = vbError + 6030
Public Const ERR_ID_EXPECTATION_STATUS_FAILED& = vbObjectError + 6500
Public Const ERR_ID_UNKNOWN_MSG_TYPE& = vbObjectError + 6600
Public Const ERR_ID_UNSUPPORTED_ARRAY_ERROR& = vbObjectError + 6700

Public Const STATUS_OK$ = "OK"
Public Const STATUS_FAIL$ = "FAIL"
Public Const STATUS_MISSING$ = "MISSING"
Public Const STATUS_PENDING$ = "PENDING"
Public Const STATUS_SKIPPED$ = "SKIPPED"


Public Const STEP_TYPE_GIVEN$ = "Given"
Public Const STEP_TYPE_WHEN$ = "When"
Public Const STEP_TYPE_THEN$ = "Then"

Public Const STEP_ATTR_TYPE$ = "step_type"
Public Const STEP_ATTR_HEAD$ = "step_head"
Public Const STEP_ATTR_NAME$ = "step_name"
Public Const STEP_ATTR_ERROR$ = "error"

Public Const REPORT_MSG_TYPE$ = "msg_type"
Public Const REPORT_MSG_CONTENT$ = "msg_content"
Public Const REPORT_MSG_STATUS$ = "msg_status"
Public Const REPORT_MSG_ERR$ = "msg_error"

Public Const REPORT_MSG_TYPE_FEATURE_NAME$ = "msg_type_feature_name"
Public Const REPORT_MSG_TYPE_DESC$ = "msg_type_description"
Public Const REPORT_MSG_TYPE_RULE$ = "msg_type_rule"
Public Const REPORT_MSG_TYPE_EXAMPLE_TITLE$ = "msg_type_example_title"
Public Const REPORT_MSG_TYPE_STEP$ = "msg_type_step"
Public Const REPORT_MSG_TYPE_CODE_TEMPLATE$ = "msg_type_code_template"
Public Const REPORT_MSG_TYPE_STATS$ = "msg_type_stats"
Public Const REPORT_MSG_TYPE_PARSE_ERR$ = "msg_type_parse_err"

Public Const LOG_TYPE_STEP$ = "log_step"
Public Const LOG_TYPE_EXAMPLE$ = "log_example"
Public Const LOG_TYPE_RULE$ = "log_rule"
Public Const LOG_TYPE_FEATURE$ = "log_feature"
