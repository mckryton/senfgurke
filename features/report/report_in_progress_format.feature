Ability: Report in progress format
    When running examples for all features, a verbose output would be confusing.
    Therefore the progress format will mark successul executed examples with
    a single dot to indicate the execution progress. Different results can be
    detected by a matching letter (e.g. F for failed steps).

  Background:
    Given the report format is "progress"

  Rule: messages about test result send to the report should result in single letters
    Most actions during test execution are resulting in messages send to the
    report. Depending on the report type those messages are formatted accordingly.

    Example: successful step
      Given a report message about a "Given a sample step" with the status "OK"
      When the reported message is formatted
      Then the resulting report output is "."

    Example: failed step
      Given a report message about a "Given a sample step" with the status "FAIL"
      When the reported message is formatted
      Then the resulting report output is "F"

    Example: missing step
      Given a report message about a "Given a sample step" with the status "MISSING"
      When the reported message is formatted
      Then the resulting report output is "M"

    Example: pending step
      Given a report message about a "Given a sample step" with the status "PENDING"
      When the reported message is formatted
      Then the resulting report output is "P"

    Example: skipped step
      Given a report message about a "Given a sample step" with the status "SKIPPED"
      When the reported message is formatted
      Then the resulting report output is "S"

    Example: three successful steps
      Given 3 steps were reported as successful
       When step results are reported
       Then the resulting report output is "..."


  Rule: code templates should reported unchanged

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
    #TODO: add feature name + example name

    Example: Step fails with an error message
      Given a report message "Given a sample step" with status "FAIL" and error msg "sample err msg"
       When the reported message is formatted
       Then the resulting report output is
        """
          F
          Err in step: Given a sample step
            sample err msg
        """
