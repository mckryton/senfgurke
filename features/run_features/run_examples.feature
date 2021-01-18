Ability: run examples
    Senfgurke will locate the matching step implementation for each step
    of an example and execute it.

  Rule: after an failed step all existing steps will return with status "skipped"

    Example: example has 3 steps and 2nd fails
       Given an example with matching step implementations
        """
          Ability: sample feature
            Example: sample example
              Given a valid step
                And an invalid step
                And a valid step
        """
       When the example is executed
       Then step statistics is "3 steps (1 passed, 1 failed, 1 skipped)"


  Rule: for given tags only examples with a matching tag are executed

    Example: feature with one matching and one non-matching example
      Given a feature has a first example "tagged example" with the tag "@sample"
      And the feature has a second example "un-tagged example" without any tag
      When the feature is executed with tag parameter "@sample"
      Then only the example "tagged example" is executed
