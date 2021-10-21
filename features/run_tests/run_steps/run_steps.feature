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


  Rule: step implementations raising an error should return "FAIL" as status

    Example: test step does not match expectation
      Given a step with a failing implementation
      When the step is executed
      Then the execution result is "FAIL"


  Rule: steps without any implementation should return "MISSING" as status

    Example: test step does not match any implemenation
      Given a step without an implementation
      When the step is executed
      Then the execution result is "MISSING"


  Rule: steps should return "PENDING" state when set explicitly
    Setting this state indicates work to be done for this step.

    Example: test step without pending message
      Given a step returning PENDING without a pending message
      When the step is executed
      Then the execution result is "PENDING"


  Rule: steps functions should be called with step expressions and data tables as parameters

    Example: calling step function for a step with a data_table
      Given a step
          """
            Given a list of numbers
              | number |
              |      3 |
              |      7 |
              |     42 |
          """
       When the step function is called
       Then the data table is passed as a parameter
