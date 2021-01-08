Ability: Report in verbose format
    While developing new features or debugging selected examples the verbose
    report format come in handy. It will show step definition next to the latest
    execution result as well as name and descriptions for features and rules.

  Background:
    Given the report format is "verbose"

  Rule: feature and rule  names and description are only shown when a following example is executed

    Example: rule without example
      Given a feature "sample feature" was reported
        And a rule "empty rule" was reported
        And a rule "has example" was reported
        And an example "sample example" was reported
       When the execution of the first example step is reported
       Then the report output contains "Feature: sample feature"
        And the report output contains "Rule: has example"
        And the report output doesn't contain "Rule: empty rule"


  Rule: verbose report will indent the feature description by one tab

    Example: feature description without leading whitespace
      Given a description line "description line"
       When the description line is formatted
       Then the description is reported as "    description line"

    Example: feature description with leading whitespace
      Given a description line "             description line"
       When the description line is formatted
       Then the description is reported as "    description line"


  Rule: format step results and names
    A step will be preceded with it's result in capital letters initially
    intendented by 5 spaces and step names following after 14 characters
    from the left. Docstrings indicators are indented by 16 spaces while
    docsting content is indented by 18 spaces.

    Example: successful step
      Given a step "Given a sample step" with the status "OK"
       When the reported message is formatted
       Then the resulting output is "     OK       Given a sample step"

    Example: failed step
      Given a step "Given a sample step" with the status "FAIL"
       When the reported message is formatted
       Then the resulting output is "     FAIL     Given a sample step"

    Example: missing step
      Given a step "Given a sample step" with the status "MISSING"
       When the reported message is formatted
       Then the resulting output is "     MISSING  Given a sample step"

    Example: pending step
      Given a step "Given a sample step" with the status "PENDING"
       When the reported message is formatted
       Then the resulting output is "     PENDING  Given a sample step"

    Example: successful step with a docstring
      Given a step "Given a sample step" followed by a docstring "this is a docstring" with the status "OK"
       When the reported message is formatted
       Then the first line of the resulting output is "     OK       Given a sample step"
        And 2nd and 4th line are "                \"\"\""
        And the 3rd line is "                  this is a docstring"


  Rule: Failed steps will show the error message after the step indented

    Example: Step fails with an error message
      Given a step "Given a sample step" fails with the error message "err: sample err msg"
       When the reported message is formatted
       Then the first line of the resulting output is "     FAIL     Given a sample step"
        And the second line shows the indented error message


  Rule: Code templates should be introduced with an explanation

    Example: Missing step
      Given a code template for a missing step was reported as
      """
Public Sub Given_a_missing_step_6A350234BFE5()
  'And a missing step
  pending
End Sub
      """
       When the report will report code templates for the missing steps
       Then the resulting output is
     """
You can implement step definitions for undefined steps with these snippets:

Public Sub Given_a_missing_step_6A350234BFE5()
 'And a missing step
 pending
End Sub
     """
