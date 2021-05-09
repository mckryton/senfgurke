Ability: run steps
    Steps are the glue that connects features with code. Every step can have a
    matching function (step implementation). Executing those functions will
    return one of the following states: OK, PENDING, MISSING or FAIL. If the
    previous step didn't return OK, all the following steps of an example will
    return SKIPPED.

  Rule: step implementations without errors should return "OK" as status

    Example: empty test step
      Given a step with an empty implementation
      When the step is executed
      Then the execution result is "OK"

    Example: test step does match expectation
      Given a step with a working implementation
      When the step is executed
      Then the execution result is "OK"


  Rule: step implementations raising an errors should return "FAIL" as status

    Example: test step does not match expectation
      Given a step with a failing implementation
      When the step is executed
      Then the execution result is "FAIL"


  Rule: steps without any implementation should return "MISSING" as status

    Example: test step does not match any implemenation
      Given a step without an implementation
      When the step is executed
      Then the execution result is "MISSING"


  Rule: steps should return "PENDING" state when ste explitcitly
    Setting this state indicates work to be done for this step.

    Example: test step without pending message
      Given a step returning PENDING without a pending message
      When the step is executed
      Then the execution result is "PENDING"


  Rule: missing steps should return a code template for the step implementation

    Example: simple step function
      Given a step "Given a valid step"
      When the code template for the step implementation is requested
      Then the code template for the step implementation is
        """
          Public Sub Given_a_valid_step_6A35DF3A18EC()
            'Given a valid step
            pending
          End Sub
        """


  Rule: suggested step function names should include the first part of the step name and a hash
    The name of the step implemenation function is the full step name stripped
    from illegal characters and spaces replaced with underscore + hash from the
    original step name. The hash is calculated using the full step name as
    input.

    Example: simple step name
      Given a step "Given a step"
      When the step is translated into a function name
      Then the name of the resulting function is "Given_a_step_C72276450E70"
