VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Steps_parse_outlines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Given_an_example_in_a_feature_file_is_defined_as_S_0AC2D1C0D26A(example_context As TContext, step_expressions As Collection)
    'Given an example in a feature file is defined as {string}
    example_context.set_value step_expressions(1), "example_as_gherkin"
End Sub

Public Sub When_the_outline_example_is_parsed_8289C92EA874(example_context As TContext)
    'When the outline example is parsed
    Dim parent_feature As TFeature
    Dim parsed_example As TExample
    Dim example_as_gherkin As String
    
    example_as_gherkin = example_context.get_value("example_as_gherkin")
    Set parent_feature = THelper.new_TFeature
    'On Error GoTo exception_handler
    Set parsed_example = TFeatureParser.parse_example(example_as_gherkin, parent_feature, New Collection)
    example_context.set_value parsed_example, "parsed_example"
    Exit Sub
    
exception_handler:
    example_context.set_value Err.description, "err_msg"
End Sub

Public Sub Then_the_outline_table_is_assigned_to_the_example_5A24E85F82D6(example_context As TContext)
    'Then the outline table is assigned to the example
    Dim parsed_example As TExample
    Dim outline_table As TDataTable
    
    Set parsed_example = example_context.get_value("parsed_example")
    TSpec.expect(parsed_example.Outlines.Count).to_be 1
    Set outline_table = parsed_example.Outlines(1)("table")
    TSpec.expect(outline_table).not_to_be_nothing
End Sub

Public Sub Then_the_outline_table_has_INT_columns_named_STR_a_ECA2472C1E76(example_context As TContext, step_expressions As Collection)
    'And the outline table has {integer} columns named {string} and {string}
    Dim parsed_example As TExample
    Dim outline_table As TDataTable
    
    Set parsed_example = example_context.get_value("parsed_example")
    Set outline_table = parsed_example.Outlines(1)("table")
    TSpec.expect(outline_table.column_names.Count).to_be step_expressions(1)
    TSpec.expect(outline_table.column_names(1)).to_be step_expressions(2)
    TSpec.expect(outline_table.column_names(2)).to_be step_expressions(3)
End Sub

Public Sub Then_the_data_of_row_INT_of_the_outline_table_is_S_E87C8EB4985B(example_context As TContext, step_expressions As Collection)
    'And the data of row {integer} of the outline table is {string}
    Dim parsed_example As TExample
    Dim outline_table As TDataTable
    Dim table_row As Collection
    Dim expected_row_data As Variant
    Dim col_index As Integer
    
    Set parsed_example = example_context.get_value("parsed_example")
    Set outline_table = parsed_example.Outlines(1)("table")
    Set table_row = outline_table.table_rows(step_expressions(1))
    expected_row_data = Split(step_expressions(2), ",")
    For col_index = 0 To UBound(expected_row_data)
        TSpec.expect(table_row(col_index + 1)).to_be expected_row_data(col_index)
    Next
End Sub

Public Sub Then_the_full_name_of_the_outline_table_is_STR_BDAC061D7B22(example_context As TContext, step_expressions As Collection)
    'And the full name of the outline table is {string}
    Dim parsed_example As TExample
    
    Set parsed_example = example_context.get_value("parsed_example")
    TSpec.expect(parsed_example.Outlines(1)("full_name")).to_be step_expressions(1)
End Sub

Public Sub Then__INT_outline_tables_are_assigned_to_the_examp_EC7DC035E3AB(example_context As TContext, step_expressions As Collection)
    'Then  {integer} outline tables are assigned to the example
    Dim parsed_example As TExample
    
    Set parsed_example = example_context.get_value("parsed_example")
    TSpec.expect(parsed_example.Outlines.Count).to_be step_expressions(1)
End Sub

Public Sub Then_the_full_name_of_outline_table_INT_is_STR_BDAC37ABF023(example_context As TContext, step_expressions As Collection)
    'And the full name of outline table {integer} is {string}
    Dim parsed_example As TExample
    Dim outline_table_name As String
    
    Set parsed_example = example_context.get_value("parsed_example")
    outline_table_name = parsed_example.Outlines(step_expressions(1))("full_name")
    TSpec.expect(outline_table_name).to_be step_expressions(2)
End Sub
