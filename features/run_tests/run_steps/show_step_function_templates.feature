@wip
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

  @vba-specific
  Rule: suggested step function names should include the first part of the step name and a hash
    The name of the step implemenation function is the full step name stripped
    from illegal characters and spaces replaced with underscore + hash from the
    original step name. The hash is calculated using the full step name as
    input.

    Example: simple step name
      Given a step "Given a step"
      When the step is translated into a function name
      Then the name of the resulting function is "Given_a_step_C72276450E70"


  Rule: steps of the same type deviating only by keyword should show only one code template

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
