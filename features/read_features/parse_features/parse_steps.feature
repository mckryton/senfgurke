Ability: parse steps
  Steps are the building blocks of examples. Every step has a matching step
  definition where the step is expressed as working code. As result, an example
  can be considered "working" if all the code from the step definition is
  executed without errors.


  Rule: steps are starting with a type

    Example: step type is "Given"
      Given a step is defined as "Given a step"
       When the step definition is parsed
       Then the type of the step is set to "Given"
        And the name of the step is "a step"

    Example: step type is "When"
      Given a step is defined as "When something happens"
       When the step definition is parsed
       Then the type of the step is set to "When"
        And the name of the step is "something happens"

    Example: step type is "Then"
      Given a step is defined as "Then some result is expected"
       When the step definition is parsed
       Then the type of the step is set to "Then"
        And the name of the step is "some result is expected"

    Example: step not starting with a step type keyword
      Given a step is defined as "a pre-condition is Given"
       When the type of the step line is evaluated
       Then the resulting line type is not "step line"


  Rule: steps starting with a synonym like "And" or "But" should use the type of previous step
    This is just for a better documentation of the step context in the reports
    but without any impact on the step execution.

    Example: "And" as synonym for "Given"
      Given a list of steps
          """
            Given x is 1
              And y is 2
          """
       When the step list is parsed
       Then the type of step 2 is set to "Given"

     Example: "But" as synonym for "When"
       Given a list of steps
           """
             When some action happens
              But it doesn't matter
           """
        When the step list is parsed
        Then the type of step 2 is set to "When"


  Rule: a list of steps is terminated by a section definition or the end of the feature file

    Example: Example following a step list
      Given a list of steps
          """
              Given one step
                And another step

            Example: sample Rule
              Given a new step
          """
       When the step list is parsed
       Then the resulting step list contains only 2 steps
        And the name of step 1 is "one step"
        And the name of step 2 is "another step"

    Example: comment line between steps
      Given a list of steps
          """
            Given one step
          # this is a comment
             When something happens
          """
       When the step list is parsed
       Then the resulting step list contains only 2 steps
