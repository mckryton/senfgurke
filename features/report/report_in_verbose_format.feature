Ability: Report in verbose format
    While developing new features or debugging selected examples the verbose
    report format come in handy. It will show step definition next to the latest
    execution result as well as name and descriptions for features and rules.

  Background:
    Given the report format is "verbose"

  Rule: feature and rule names should only appear before an example

    Example: rule without example
      Given a feature "sample feature" was reported
        And a rule "empty rule" was reported
        And a rule "has example" was reported
        And an example "sample example" was reported
       When the execution of the first example step is reported
       Then the report output contains "Feature: sample feature"
        And the report output contains "Rule: has example"
        And the report output doesn't contain "Rule: empty rule"


  Rule: feature descriptions should be indented with a fixed space

    Example: feature description without leading whitespace
      Given a description line "description line"
       When the description line is formatted
       Then the description is reported as "    description line"

    Example: feature description with leading whitespace
      Given a description line "             description line"
       When the description line is formatted
       Then the description is reported as "    description line"


  Rule: format step results and names
    Most actions during test execution are resulting in messages send to the
    report. Depending on the report type those messages are formatted
    accordingly.
    For the verbose report a step will be preceded with it's result in capital
    letters initially intendented by 5 spaces and step names following after
    14 characters from the left. Docstrings indicators are indented by 16 spaces
    while docstring content is indented by 18 spaces.

    Example: successful step
      Given a report message about a "Given a sample step" with the status "OK"
       When the reported message is formatted
       Then the resulting report output is "     OK       Given a sample step"

    Example: failed step
      Given a report message about a "Given a sample step" with the status "FAIL"
       When the reported message is formatted
       Then the resulting report output is "     FAIL     Given a sample step"

    Example: missing step
      Given a report message about a "Given a sample step" with the status "MISSING"
       When the reported message is formatted
       Then the resulting report output is "     MISSING  Given a sample step"

    Example: pending step
      Given a report message about a "Given a sample step" with the status "PENDING"
       When the reported message is formatted
       Then the resulting report output is "     PENDING  Given a sample step"

    Example: successful step with a docstring
      Given a report message "Given a sample step" with the status "OK" followed by a docstring "this is a docstring"
       When the reported message is formatted
       Then line 1 of the resulting output is "     OK       Given a sample step"
        And line 2 of the resulting output is "                  this is a docstring"


  Rule: Failed steps will show the error message after the step indented

    Example: Step fails with an error message
      Given a report message "Given a sample step" with status "FAIL" and error msg "sample err msg"
       When the reported message is formatted
       Then line 1 of the resulting output is "     FAIL     Given a sample step"
        And line 2 of the resulting output is "          sample err msg"


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
       Then the resulting report output is
     """
You can implement step definitions for undefined steps with these snippets:

Public Sub Given_a_missing_step_6A350234BFE5()
 'And a missing step
 pending
End Sub
     """

     #Todo: add example for multiple missing steps in more than one scenario
