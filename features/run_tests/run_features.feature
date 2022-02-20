Ability: run features
    A features represents the business intention on a more abstract level. The
    feature description should tell you why the feature was implemented in the
    first place.
    Senfgurke runs features by finding all the examples (see the run examples
    feature for more) and executes all the steps from those examples (see the
    run steps feature for more).

  Rule: steps from background will be executed before each example

    Example: background and two examples
      Given a feature
       """
        Feature: feature has background
          Background:
            Given a background step

          Example: sample 1
            Given a step for scenario 1

          Example: sample 2
            Given another step for scenario 2
       """
      When the feature is executed
      Then the background step is executed once before each example


  Rule: for any given tag execute only those examples assigned with the tag

    Example: tag assigned to an example
      Given a feature
        """
          Feature: has example with tags
            @sample
            Example: tagged example
              Given one step

            Example: hasn't tags
              Given another step
        """
       When the feature is executed with tag parameter "@sample"
       Then only the example "tagged example" is executed

    Example: example inherits tag from feature

    Example: example inherits tag from rule
