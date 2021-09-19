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


  Rule: allow "And" and "But" as synonyms for "Given", "When", "Then"
    Translate "And" and "But" to "Given", "When" or "Then" using the same keyword
    as the previous step.

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
