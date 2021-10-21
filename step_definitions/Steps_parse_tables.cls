VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_parse_tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_a_line_STR_in_a_feature_file_A1515550652D(example_context As TContext, step_expressions As Collection)
    'Given a line {string} in a feature file
    example_context.set_value step_expressions(1), "gherkin_text"
End Sub

Public Sub Given_a_step_in_a_feature_file_is_defined_as_STR_9DBCD4C78C79(example_context As TContext, step_expressions As Collection)
    'Given a step in a feature file is defined as {string}
    example_context.set_value step_expressions(1), "gherkin_text"
End Sub

Public Sub Given_a_table_in_a_feature_file_is_defined_as_STR_FD4681DF4049(example_context As TContext, step_expressions As Collection)
    'Given a table in a feature file is defined as {string}
    example_context.set_value step_expressions(1), "gherkin_text"
End Sub

Public Sub When_the_type_of_this_line_is_parsed_4537958B6F8D(example_context As TContext)
    'When the type of this line is parsed
    Dim parent_feature As TFeature
    
    Set parent_feature = New TFeature
    parent_feature.parsed_lines = -1
    On Error GoTo exception_handler
    example_context.set_value TFeatureParser.get_line_type(example_context.get_value("gherkin_text")), "line_type"
    example_context.set_value TStepParser.parse_table(example_context.get_value("gherkin_text"), parent_feature), "data_table"
    Exit Sub
    
exception_handler:
    example_context.set_value Err.description, "err_msg"
End Sub

Public Sub When_the_table_is_parsed_CEB05834DFDE(example_context As TContext)
    'When the table is parsed
    Dim parent_feature As TFeature
    
    Set parent_feature = New TFeature
    parent_feature.parsed_lines = -1
    On Error GoTo exception_handler
    example_context.set_value TStepParser.parse_table(example_context.get_value("gherkin_text"), parent_feature), "parsed_table"
    Exit Sub
    
exception_handler:
    example_context.set_value Err.description, "err_msg"
End Sub

Public Sub Then_the_line_is_recognized_as_part_of_a_table_ADEDA52A5184(example_context As TContext)
    'Then the line is recognized as part of a table
    TSpec.expect(example_context.get_value("line_type")).to_be LINE_TYPE_TABLE_ROW
End Sub

Public Sub Then_the_table_row_has_one_column_398C1CB663F5(example_context As TContext)
    'And the table row has one column
    Dim data_table As TDataTable
    
    Set data_table = example_context.get_value("data_table")
    TSpec.expect(data_table.column_names.Count).to_be 1
End Sub

Public Sub Then_the_table_row_has_INT_columns_398C1CB6110F(example_context As TContext, step_expressions As Collection)
    'And the table row has {integer} columns
    Dim data_table As TDataTable
    
    If example_context.value_exists("data_table") Then
        Set data_table = example_context.get_value("data_table")
    Else
        'pick the data table form the first step ofa step list
        Set data_table = example_context.get_value("step_list")(1).data_table
        example_context.set_value data_table, "data_table"
    End If
    TSpec.expect(data_table.column_names.Count).to_be step_expressions(1)
End Sub

Public Sub Then_parsing_will_return_the_syntax_error_STR_04E34D1BD0C6(example_context As TContext, step_expressions As Collection)
    'Then parsing will return the syntax error {string}
    TSpec.expect(example_context.get_value("err_msg")).to_be step_expressions(1)
End Sub

Public Sub Then_the_line_is_recognized_as_description_6CFBBCC51217(example_context As TContext)
    'Then the line is recognized as description
    TSpec.expect(example_context.get_value("line_type")).to_be LINE_TYPE_DESCRIPTION
End Sub

Public Sub Then_the_data_table_is_assigned_to_the_step_CA3B18FB49A3(example_context As TContext)
    'Then the data table is assigned to the step
    Dim step_list As Collection
    Dim step As TStep
    
    Set step_list = example_context.get_value("step_list")
    Set step = step_list(1)
    TSpec.expect(step.data_table).not_to_be_nothing
End Sub

Public Sub Then_the_names_of_the_columns_are_STR_and_STR_1C91C352E272(example_context As TContext, step_expressions As Collection)
    'And the names of the columns are {string} and {string}
    Dim data_table As TDataTable
    
    Set data_table = example_context.get_value("data_table")
    
    TSpec.expect(data_table.column_names).contains_member step_expressions(1)
    TSpec.expect(data_table.column_names).contains_member step_expressions(2)
End Sub

Public Sub Then_the_table_contains_only_string_values_4E9F8D8EFD50(example_context As TContext)
    'Then the table contains only string values
    Dim parsed_table As TDataTable
    Dim table_row As Variant
    
    Set parsed_table = example_context.get_value("parsed_table")
    For Each table_row In parsed_table.table_rows
        TSpec.expect(TypeName(table_row("text"))).to_be "String"
    Next
End Sub

Public Sub Then_the_first_table_row_doesnt_contain_quotes_B5958A779CAE(example_context As TContext)
    'And the first table row doesn't contain quotes
    Dim parsed_table As TDataTable
    
    Set parsed_table = example_context.get_value("parsed_table")
    TSpec.expect(parsed_table.table_rows(1)("text")).does_not_include """"
End Sub
