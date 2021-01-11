Ability: Report in progress format
    When running examples form all features a verbose output would be confusing.
    Therefore the progress format will mark successul executed examples with
    a single dot to indicate the execution progress. Different results can be
    detected by a matching letter (e.g. F for failed steps).

  Background:
    Given the report format is "progress"

  Rule: format step results and names

    Example: successful step
      Given a step "Given a sample step" with the status "OK"
      When the reported message is formatted
      Then the resulting report output is set to "."

    Example: failed step
      Given a step "Given a sample step" with the status "FAIL"
      When the reported message is formatted
      Then the resulting report output is set to "F"

    Example: missing step
      Given a step "Given a sample step" with the status "MISSING"
      When the reported message is formatted
      Then the resulting report output is set to "M"

    Example: pending step
      Given a step "Given a sample step" with the status "PENDING"
      When the reported message is formatted
      Then the resulting report output is set to "P"

    Example: three successful steps
      Given 3 steps were reported as successful
       When step results are reported
       Then the resulting report output is "..."


  Rule: Keep code templates as is

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


  Rule: Failed steps should show step name and error message

    Example: Step fails with an error message
      Given a step "Given a sample step" fails with the error message "err: sample err msg"
        And the report format is set to progress
       When the reported message is formatted
       Then the first line of the resulting output is set to "F"
        And the second line is "Err in step: Given a sample step"
        And the third line is "  sample err msg"
