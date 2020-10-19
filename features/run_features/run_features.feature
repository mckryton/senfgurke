Ability: run features
    Senfgurke will locate the matching step implementation for each step
    in a feature and execute it.

  Rule: steps from background will be executed before each example

    Example: background and two examples
      Given a feature has a background step
      And the feature has two examples with on step
      When the feature is executed
      Then the background step is executed once before each example
