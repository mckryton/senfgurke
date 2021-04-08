Ability: run features
    A features represents the business intention on a more abstract level. The
    feature description should tell you why the feature was implemented in the
    first place.
    Senfgurke runs feature by finding all the examples (see run examples
    feature for more) and executing all the steps from those examples (see run
    steps feature for more).

  Rule: steps from background will be executed before each example

    Example: background and two examples
      Given a feature has a background step
      And the feature has two examples with on step
      When the feature is executed
      Then the background step is executed once before each example
