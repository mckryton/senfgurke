Ability: Report in verbose format
    While executing examples senfgurke will send messages about progress
    and success for reporting. The verbose formatter will turn those messages into
     a verbose report that is printed on the debug console.


  Rule: feature and rule  names and description are only shown when a following example is executed

    Example: rule without example
      Given the report format is set to verbose
      And a feature "sample feature" was reported
      And a rule "empty rule" was reported
      And a rule "has example" was reported
      And an example "sample example" was reported
      When the execution of the first example step is reported
      Then the report output contains "Feature: sample feature"
      And the report output contains "Rule: has example"
      And the report output doesn't contain "Rule: empty rule"


  Rule: verbose report will indent the feature description by one tab

    Example: feature description without leading whitespace
      Given the report format is set to verbose
      And a line from the feature is reported as feature description "description line"
      When the reported message is prepared as output for the report
      Then the description is reported as "    description line"

    Example: feature description with leading whitespace
      Given the report format is set to verbose
      And a feature with a description "             description line"
      When the reported message is prepared as output for the report
      Then the description is reported as "    description line"


  Rule: A step will be preceded with it's result in capital letters initially intendented by 5 spaces and step names following after 14 characters from the left

    Example: successful step
      Given a step "Given a sample step" with the status "OK"
      And the report format is set to verbose
      When the reported message is prepared as output for the report
      Then the resulting output is set to "     OK       Given a sample step"

    Example: failed step
      Given a step "Given a sample step" with the status "FAIL"
      And the report format is set to verbose
      When the reported message is prepared as output for the report
      Then the resulting output is set to "     FAIL     Given a sample step"

    Example: missing step
      Given a step "Given a sample step" with the status "MISSING"
      And the report format is set to verbose
      When the reported message is prepared as output for the report
      Then the resulting output is set to "     MISSING  Given a sample step"

    Example: pending step
      Given a step "Given a sample step" with the status "PENDING"
      And the report format is set to verbose
      When the reported message is prepared as output for the report
      Then the resulting output is set to "     PENDING  Given a sample step"


  Rule: Rules and example titles will add an empty line before them and indent following lines with 3 tabs

    Example: Rule over two lines
      Given a rule "this is a sample rule<br>and this is line 2 of the same rule"
      And the report format is set to verbose
      When the reported message is prepared as output for the report
      Then the resulting output starts with a new line
      And every line after the first one is indented by 7 spaces


  Rule: Failed steps will show the error message after the step indented

    Example: Step fails with an error message
      Given a step "Given a sample step" fails with the error message "err: sample err msg"
      And the report format is set to verbose
      When the reported message is prepared as output for the report
      Then the resulting output starts with the step status and name
      And the second line shows the indented error message
