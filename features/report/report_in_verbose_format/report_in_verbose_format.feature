Ability: Report in verbose format
    While developing new features or debugging selected examples the verbose
    report format come in handy. It will show step definition next to the latest
    execution result as well as name and descriptions for features and rules.

  Background:
    Given the report format is "verbose"

  Rule: feature and rule names should only appear when attached examples were executed

    Example: skip empty rule
      When the following events are reported as a result of running a feature
          | section_type | section_name     | section_status |
          | Feature      | sample feature   |                |
          | Rule         | empty rule       |                |
          | Rule         | rule has example |                |
          | Example      | sample           |                |
          | Step         | Given one step   | OK             |
      Then the resulting report output is
        """
          Feature: sample feature
            Rule: rule has example
              Example: sample
               OK       Given one step
        """

    Example: skip empty feature
      When the following events are reported as a result of running a feature
          | section_type | section_name     | section_status |
          | Feature      | empty feature    |                |
          | Feature      | filled feature   |                |
          | Rule         | rule has example |                |
          | Example      | sample           |                |
          | Step         | Given one step   | OK             |
      Then the resulting report output is
        """
          Feature: filled feature
            Rule: rule has example
              Example: sample
               OK       Given one step
        """


  Rule: descriptions shouldn't appear in verbose reports
    Verbose reports are provided for detailed feedback during devlopemnt and not
    for living documentation. To keep the feedback comprehensible any
    description text will be left out.

    Example: feature with example has description
      Given a feature
        """
          Feature: sample
            This feature has some description.

            Example: sample
              Given one step
        """
      When the execution of the feature is reported
       And the example is reported with the step result "OK"
      Then the resulting report output is
        """
          Feature: sample
              Example: sample
               OK       Given one step
        """


  Rule: executed steps should be reported with the exceution status
    Executing a step is done by calling the matching step function and returning
    the execution status as result.
    For the verbose report a step will be preceded with it's result in capital
    letters initially intendented by 5 spaces and step names following after
    14 characters from the left. Docstrings indicators are indented by 16 spaces
    while docstring content is indented by 18 spaces.

    Example: successful step
       When a step "Given a sample step" is reported with status "OK"
       Then the trimmed report output is "     OK       Given a sample step"

    Example: failed step
       When a step "Given a sample step" is reported with status "FAIL"
       Then the trimmed report output is "     FAIL     Given a sample step"

    Example: missing step
       When a step "Given a sample step" is reported with status "MISSING"
       Then the trimmed report output is "     MISSING  Given a sample step"

    Example: pending step
       When a step "Given a sample step" is reported with status "PENDING"
       Then the trimmed report output is "     PENDING  Given a sample step"

    Example: skipped step
       When a step "Given a sample step" is reported with status "SKIPPED"
       Then the trimmed report output is "     SKIPPED  Given a sample step"

    Example: successful step with a docstring
       When a step "Given a sample step" followed by a docstring "this is a docstring" is reported with status "OK"
       Then line 1 of the resulting output is "     OK       Given a sample step"
        And line 3 of the resulting output is "                  this is a docstring"

    Example: successful step with a data table
       When the following step is reported with the status "OK"
        """
          Given a step with a data table
            | column_name |
            | row 1       |
            | row 2       |
        """
       Then the report output is
        """
          OK       Given a step with a data table
                     | column_name |
                     | row 1       |
                     | row 2       |
        """


  Rule: Failed steps will show the error message after the step indented

    Example: Step fails with an error message
       When a step "Given a sample step" is reported with status "FAIL" and error msg "sample err msg"
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

          If the function is already in place, the cause for this message could be:
          * the step implementation class is not registered in the TStepRegister module
          * the code inside the step function tries to access a non-existing method or property
        """

     #Todo: add example for multiple missing steps in more than one scenario


  Rule: Parse errors should be reported as-is

    Example: report parse error in verbose report
      Given a parse error "Feature lacks feature keyword at the beginning" was found in "sample.feature"
       When the parse error is reported
       Then the resulting report output is
         """
          Error: Found syntax error while parsing feature 'sample.feature'
            Feature lacks feature keyword at the beginning
         """
