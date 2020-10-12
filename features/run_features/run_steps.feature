Ability: run steps
    Senfgurke will locate the matching step implementation for each step
    of an example and execute it.


  Rule: step implementations without errors return "OK" as status

    Example: Empty test step
      Given a step with an empty implementation
      When the step is executed
      Then the execution result is "OK"

    Example: test step does match expectation
      Given a step with a working implementation
      When the step is executed
      Then the execution result is "OK"


  Rule: step implementations raising an errors return "FAIL" as status

    Example: test step does not match expectation
      Given a step with a failing implementation
      When the step is executed
      Then the execution result is "FAIL"


  Rule: steps without any implementation return "MISSING" as status

    Example: test step does not match any implemenation
      Given a step without an implementation
      When the step is executed
      Then the execution result is "MISSING"


  Rule: missing steps return a code snippet for the step implementation

    Example: simple step function
      Given a step "Given a valid step"
      When the code template for the step implementation is requested
      Then the code template for the step implementation is "Public Sub Given_a_valid_step_8A74152FD2F9()<br>    'Given a valid step<br><br>End Sub"


  Rule: the name of the step implemenation function is the full step name  stripped from illegal characters and spaces replaced with underscore + hash from the original step name

    Example: simple step name
      Given a step "Given a step"
      When the step is translated into a function name
      Then the name of the resulting function is "Given_a_step_C7224E350E70"
