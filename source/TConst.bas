Attribute VB_Name = "TConst"
Option Explicit

' see TError module for error related constants

Public Const SECTION_TYPE_FEATURE$ = "feature"
Public Const SECTION_TYPE_RULE$ = "rule"
Public Const SECTION_TYPE_EXAMPLE$ = "example"
Public Const SECTION_TYPE_BACKGROUND$ = "background"

Public Const SECTION_ATTR_TYPE$ = "type"
Public Const SECTION_ATTR_NAME$ = "name"

Public Const STATUS_OK$ = "OK"
Public Const STATUS_FAIL$ = "FAIL"
Public Const STATUS_MISSING$ = "MISSING"
Public Const STATUS_PENDING$ = "PENDING"
Public Const STATUS_SKIPPED$ = "SKIPPED"
Public Const STATUS_UNKNOWN$ = "UNKNOWN"


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

Public Const LOG_EVENT_RUN_STEP$ = "log_run_step"
Public Const LOG_EVENT_RUN_EXAMPLE$ = "log_run_example"
Public Const LOG_EVENT_RUN_RULE$ = "log_run_rule"
Public Const LOG_EVENT_RUN_FEATURE$ = "log_run_feature"
Public Const LOG_EVENT_LOAD_FEATURE$ = "log_load_feature"


'------- Parser Constants -------
Public Const START_SECTION_KEYWORDS$ = "Example,Scenario,Scenario Outline,Rule,Background,Ability,Business Needs,Feature"

Public Const LINE_TYPE_FEATURE_START$ = "feature start"
Public Const LINE_TYPE_RULE_START$ = "rule start"
Public Const LINE_TYPE_EXAMPLE_START$ = "example start"
Public Const LINE_TYPE_BACKGROUND_START$ = "background start"
Public Const LINE_TYPE_TAGS$ = "tag line"
Public Const LINE_TYPE_STEP$ = "step line"
Public Const LINE_TYPE_DESCRIPTION$ = "description line"
Public Const LINE_TYPE_COMMENT$ = "comment line"
Public Const LINE_TYPE_DOCSTRING_LIMIT$ = "docstring limit"
Public Const LINE_TYPE_TABLE_ROW$ = "table row"

