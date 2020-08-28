Ability: show verbose results
    While executing examples senfgurke will send messages about progress
    and success for reporting. The verbose formatter will turn those messages into
     a verbose report that is printed on the debug console.

  # Rule: verbose report will indent the feature description by one tab

  Example: feature description with 2 lines
     Given a feature with a description "description line 1 <newline> description line 2"
     And the report format is set to verbose
     When the reported message is prepared as output for the report
     Then every line of the description is indented by one tab

  Example: feature description with a single line
     Given a feature with a description "description line 1"
     And the report format is set to verbose
     When the reported message is prepared as output for the report
     Then every line of the description is indented by one tab

  # Rule: A step will be preceded with it's result in capital letters initially intendented by 1 tab
  #       and then indented by 1 space and steps following after indented by 10 spaces

  Example: successful step
     Given a step "this is a step" with the status "OK"
     And the report format is set to verbose
     When the reported message is prepared as output for the report
     Then the resulting output is set to tab and "   this is a step"

  Example: failed step
     Given a step "this is a step" with the status "FAILED"
     And the report format is set to verbose
     When the reported message is prepared as output for the report
     Then the resulting output is set to tab and " FAILED   this is a step"

  Example: missing step
     Given a step "this is a step" with the status "MISSING"
     And the report format is set to verbose
     When the reported message is prepared as output for the report
     Then the resulting output is set to tab and " MISSING  this is a step"

  # Rule: Rules and example titles will add an empty line before them
  #         and indent following lines with 3 tabs

  Example: Rule over two lines
     Given a rule "this is a sample rule<br>and this is line 2 of the same rule"
     And the report format is set to verbose
     When the reported message is prepared as output for the report
     Then the resulting output starts with a new line
     And every line after the first one is indented by 7 spaces

  # Rule: Failed steps will show the error message after the step indented

  Example: Step fails with an error message
     Given a step "this step fails" with status "FAILED" and with the error message "err: sample err msg"
     And the report format is set to verbose
     When the reported message is prepared as output for the report
     Then the resulting output starts with the step status and description
     And the second line shows the indented error message
