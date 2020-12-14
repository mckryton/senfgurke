Ability: parse examples
  Examples (aka scenarios) are the executable part of a feature. An example
  will explaining a rule in three steps: pre-conditions, action and results.
  Parsing examples will translate plain text into structures that set the right
  relation between examples and rules as well as steps and examples.

  Rule: ignore leading whitespace befor steps
    Every line in a example clause starting with
    "<optional whitespace><step keyword><space>" is considered to be a step
    where <tep keyword> is one of those: Given, When, Then, And, But

    Example: simple example with Given, When, Then steps
      Given an example "sample example"
      And the example contains a step "    Given a precondition"
      And the example contains a step "    When action happens"
      And the example contains a step "    Then some result is expected"
      When the example is parsed
      Then the example clause from the parsed result contains all the steps

#    Rule: Steps not starting with Given, When, Then or And will produce a Gherkin syntax error
#    Example: missing step type
#      Given an example with a step "bla bla bla"
#      When the step is read
#      Then the step result is an Gherkin Syntax Error claiming "bla" is not a valid step type
