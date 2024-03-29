Attribute VB_Name = "TConst"
Option Explicit

' see TError module for error related constants

Public Const SECTION_TYPE_FEATURE$ = "feature"
Public Const SECTION_TYPE_RULE$ = "rule"
Public Const SECTION_TYPE_EXAMPLE$ = "example"
Public Const SECTION_TYPE_BACKGROUND$ = "background"
Public Const SECTION_TYPE_STEP$ = "step"
Public Const SECTION_TYPE_OUTLINE$ = "outline"

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
Public Const REPORT_MSG_FEATURE_ORIGIN$ = "msg_feature_origin"

Public Const REPORT_MSG_TYPE_FEATURE_NAME$ = "msg_type_feature_name"
Public Const REPORT_MSG_TYPE_DESC$ = "msg_type_description"
Public Const REPORT_MSG_TYPE_RULE$ = "msg_type_rule"
Public Const REPORT_MSG_TYPE_EXAMPLE_TITLE$ = "msg_type_example_title"
Public Const REPORT_MSG_TYPE_STEP$ = "msg_type_step"
Public Const REPORT_MSG_TYPE_OUTLINE_STEP$ = "msg_type_outline_step"
Public Const REPORT_MSG_TYPE_CODE_TEMPLATE$ = "msg_type_code_template"
Public Const REPORT_MSG_TYPE_STATS$ = "msg_type_stats"
Public Const REPORT_MSG_TYPE_PARSE_ERR$ = "msg_type_parse_err"
Public Const REPORT_MSG_TYPE_OUTLINE_TABLEHEADER$ = "msg_type_outline_table_header"
Public Const REPORT_MSG_TYPE_OUTLINE_ROW$ = "msg_type_outline_row"

Public Const EVENT_FEATURE_LOADED$ = "feature_loaded"
Public Const EVENT_PARSE_ERROR = "parse_error"
Public Const EVENT_RUN_SESSION_FINISHED$ = "run_session_finished"
Public Const EVENT_RUN_FEATURE_STARTED$ = "run_feature_started"
Public Const EVENT_RUN_FEATURE_FINISHED$ = "run_feature_finished"
Public Const EVENT_RUN_RULE_STARTED$ = "run_rule_started"
Public Const EVENT_RUN_EXAMPLE_STARTED$ = "run_example_started"
Public Const EVENT_RUN_EXAMPLE_FINISHED$ = "run_example_finished"
Public Const EVENT_RUN_OUTLINE_EXAMPLE_STARTED$ = "run_outline_example_started"
Public Const EVENT_RUN_OUTLINE_EXAMPLE_FINISHED$ = "run_outline_example_finished"
Public Const EVENT_RUN_OUTLINE_TABLE_STARTED$ = "run_outline_table_started"
Public Const EVENT_RUN_OUTLINE_ROW_STARTED$ = "run_outline_row_started"
Public Const EVENT_RUN_STEP_FINISHED$ = "run_step_finished"
Public Const EVENT_RUN_OUTLINE_STEP_DECLARED$ = "run_outline_step_declared"
Public Const EVENT_STEP_IS_MISSING = "step_is_missing"


'------- Parser Constants -------
Public Const START_SECTION_KEYWORDS$ = "Example,Scenario,Scenario Outline,Rule,Background,Ability,Business Needs,Feature,Examples"

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
Public Const LINE_TYPE_OUTLINE_START$ = "outline table"

