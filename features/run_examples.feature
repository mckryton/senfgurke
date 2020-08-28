Ability: run examples
    Senfgurke will validate the Gherkin syntax of the given examples
    and calculate a result for every step of the example.

  # Rule: implemented test steps are OK if no error was raised, otherwise the steps FAIL

  Example: Empty test step
     Given an example with a step "Given a step with an empty implementation"
     And the step has a matching step implementation
     When the step is executed
     Then the execution result is "OK"

  Example: test step does match expectation
     Given an example with a step "Given a step with a valid expectation"
     And the step has a matching step implementation
     When the step is executed
     Then the execution result is "OK"

  Example: test step does not match expectation
     Given an example with a step "Given a step with an invalid expectation"
     And the step has a matching step implementation
     When the step is executed
     Then the execution result is "FAIL"
