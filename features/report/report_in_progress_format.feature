Ability: Report in progress format
  Running examples for all features would be confusing because of the amount
  of details for all the steps.
  Therefore the progress format will mark successul executed examples with
  a single dot to indicate the execution progress. Different results can be
  detected by a matching letter (e.g. F for failed steps).

  Background:
    Given the report format is "progress"

  Rule: messages about test result send to the report should result in single letters
    Most actions during test execution are resulting in messages send to the
    report. Depending on the report type those messages are formatted accordingly.

    Example: successful step
       When a step "Given a sample step" is reported with status "OK"
       Then the resulting report output is "."

    Example: failed step
       When a step "Given a sample step" is reported with status "FAIL"
       Then the resulting report output is "F"

    Example: missing step
       When a step "Given a sample step" is reported with status "MISSING"
       Then the resulting report output is "M"

    Example: pending step
       When a step "Given a sample step" is reported with status "PENDING"
       Then the resulting report output is "P"

    Example: skipped step
       When a step "Given a sample step" is reported with status "SKIPPED"
       Then the resulting report output is "S"

    Example: three successful steps
       When 3 steps are reported as successful
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

          If the function is already in place, the cause for this message could be:
          * the step implementation class is not registered in the TStepRegister module
          * the code inside the step function tries to access a non-existing method or property
        """


  Rule: Failed steps should show step name and error message

    Example: single failed step
      Given the origin for some steps is "sample.feature"
        And steps were reported as
          | step_name             | status | err_msg        |
          | Given an invalid step | FAIL   | sample err msg |
       When all steps were reported and the report is finished
       Then the resulting report output is
        """
          F
          Err in step: Given an invalid step
          (feature: sample.feature)
            sample err msg
        """

    Example: failed step followed by passed step
      Given the origin for some steps is "sample.feature"
        And steps were reported as
          | step_name             | status | err_msg        |
          | Given an invalid step | FAIL   | sample err msg |
          | And a valid step      | OK     |                |
       When all steps were reported and the report is finished
       Then the resulting report output is
        """
          F.
          Err in step: Given an invalid step
          (feature: sample.feature)
            sample err msg
        """

  Rule: verbose formatter should add a line break before every 81st step

    Example: report 85 steps
       When 85 steps are reported as successful
       Then the resulting report output is
        """
          ................................................................................
          .....
        """


  Rule: Parse errors should be reported as-is

    Example: report parse error in progress report
     Given a parse error "Feature lacks feature keyword at the beginning" was found in "sample.feature"
      When the parse error is reported
      Then the resulting report output is
        """
         Error: Found syntax error while parsing feature 'sample.feature'
           Feature lacks feature keyword at the beginning
        """
