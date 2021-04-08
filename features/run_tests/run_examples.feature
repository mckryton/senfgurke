Ability: run examples
    Examples are the building blocks for describing the application under test.
    Running examples means that Senfgurke will try to find and execute a
    matching function for each step in an example (aka step implementation).
    The result of the executed step will tell you if the application works as
    expected (see feature "run steps" for more).

  Rule: steps follwing a step with a status other than "OK" should return "SKIPPED" or "MISSING"

    Example: example with failed step
       Given an example with matching step implementations
        """
          Ability: sample feature
            Example: sample example
              Given a valid step
                And an invalid step
                And a valid step
                And a valid step
        """
       When the example is executed
       Then step statistics is "4 steps (1 passed, 1 failed, 2 skipped)"

     Example: example with pending step
        Given an example with matching step implementations
         """
           Ability: sample feature
             Example: sample example
               Given a valid step
                 And a pending step
                 And a valid step
                 And a valid step
         """
        When the example is executed
        Then step statistics is "4 steps (1 passed, 1 pending, 2 skipped)"

  #TODO: add example for missing steps

  # to detect missing steps, steps have to be executed even when they are skipped
  #TODO: add rule not to show err msg for skipped steps

  Rule: for given tags only examples with a matching tag are executed

    Example: feature with one matching and one non-matching example
      Given a feature has a first example "tagged example" with the tag "@sample"
      And the feature has a second example "un-tagged example" without any tag
      When the feature is executed with tag parameter "@sample"
      Then only the example "tagged example" is executed
