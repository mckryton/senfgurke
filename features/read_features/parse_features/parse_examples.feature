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
      Given a feature
      And the feature includes a line "  Example: sample example"
      And this line is followed by "    Given a precondition"
      And this line is followed by "    When action happens"
      And this line is followed by "    Then some result is expected"
      When the feature is parsed
      Then the parsed result contains an example
      And the example clause from the parsed result contains all the steps


  Rule: allow And and But as synonyms for Given, When, Then

          Translate "And" and "But" to Given, When or Then using the same keyword
          as the previous step.

    Example: And as synonym for Given
      Given an example with two steps "Given x is 1" and "And y is 2"
      When the feature is parsed
      Then the type of the second step is set to "Given"

    Example: But as synonym for When
      Given an example with two steps "When some action happens" and "But it doesn't matter"
      When the feature is parsed
      Then the type of the second step is set to "When"


  Rule: add docstrings to the previous step

      A docstring is a multiline string that is related to the previous step.
      Docstrings are embraced by a sequence of 3 double quotation marks """

      Example: example with docstring
        Given an example
          And the first step is "Given a first step"
          And this step is followed by a docstring containing "this is a docstring"
         When the example is parsed
         Then first step is changed to "Given a first step \"this is a docstring\""

      Example: feature background with docstring
        Given a background
          And the first step of the background is "Given a first background step"
          And this step is followed by a docstring containing "this is a docstring"
         When the feature background is parsed
         Then first step of the background is changed to "Given a first background step \"this is a docstring\""


#    Rule: Steps not starting with Given, When, Then or And will produce a Gherkin syntax error
#    Example: missing step type
#      Given an example with a step "bla bla bla"
#      When the step is read
#      Then the step result is an Gherkin Syntax Error claiming "bla" is not a valid step type
