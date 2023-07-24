Ability: show step function templates
  Whenever a step is executed and no matching step function is found Senfgurke
  will generate a code snippet for a new step function. This way it should be
  easy to assign code to the given step.


  Rule: missing steps should return a code template for the step implementation

    Example: step function template without step expressions
      Given a step "Given a valid step"
       When the code template for the step implementation is requested
       Then the code template for the step implementation is
          """
            Public Sub Given_a_valid_step_6A35DF3A18EC(example_context as TContext)
                'Given a valid step
                pending
            End Sub
          """

    Example: step function template with step expressions
      Given a step "Given a number 42"
       When the code template for the step implementation is requested
       Then the code template for the step implementation is
          """
            Public Sub Given_a_number_INT_6A352C3C78C6(example_context as TContext, step_expressions As Collection)
                'Given a number {integer}
                pending
            End Sub
          """

    Example: step function template with data table
      Given a step
          """
            Given a list of numbers
              | number |
              |      3 |
              |      7 |
              |     42 |
          """
       When the code template for the step implementation is requested
       Then the code template for the step implementation is
          """
            Public Sub Given_a_list_of_numbers_FF6B22C2A68F(example_context as TContext, data_table as TDataTable)
                'Given a list of numbers
                pending
            End Sub
          """

    Example: step function template with step expressions and data table
      Given a step
          """
            Given 42 in a list of numbers
              | number |
              |      3 |
              |      7 |
              |     42 |
          """
       When the code template for the step implementation is requested
       Then the code template for the step implementation is
          """
            Public Sub Given__INT_in_a_list_of_numbers_4DD8A2DD2981(example_context as TContext, step_expressions As Collection, data_table as TDataTable)
                'Given  {integer} in a list of numbers
                pending
            End Sub
          """

  Rule: steps using a synonym keyword should show only one code template

    Example: Two steps using Given and And
      Given an example
        """
          Feature: sample
            Example: two steps
              Given a new step
                And a new step
        """
        And the report format is "verbose"
       When the example is executed
       Then the resulting report will show the code template only once
