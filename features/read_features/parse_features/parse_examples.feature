Ability: parse examples
  Senfgurke will read feature specs from text
  and identify examples and their example steps from within features


  Rule: every line in a example clause starting with "<optional whitespace><step keyword><space>" is considered to be a step where <tep keyword> is one of those: Given, When, Then, And, But

    Example: simple example with Given, When, Then steps
      Given a feature
      And the feature includes a line "  Example: sample example"
      And this line is followed by "    Given a precondition"
      And this line is followed by "    When action happens"
      And this line is followed by "    Then some result is expected"
      When the feature is parsed
      Then the parsed result contains an example
      And the example clause from the parsed result contains all the steps


  Rule: Translate "And" and "But" to Given, When or Then using the same keyword as the previous step

    Example: And as synonym for Given
      Given an example with two steps "Given x is 1" and "And y is 2"
      When the feature is parsed
      Then the type of the second step is set to "Given"

    Example: But as synonym for When
      Given an example with two steps "When some action happens" and "But it doesn't matter"
      When the feature is parsed
      Then the type of the second step is set to "When"

#    Rule: Steps not starting with Given, When, Then or And will produce a Gherkin syntax error
#    Example: missing step type
#      Given an example with a step "bla bla bla"
#      When the step is read
#      Then the step result is an Gherkin Syntax Error claiming "bla" is not a valid step type


# TODO: warn when steps ar to long (255/63 chars)
