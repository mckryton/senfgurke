Ability: run examples
    Senfgurke will locate the matching step implementation for each step
    of an example and execute it.

  Rule: examples are executed until a step returns a status other than "OK"

  Example: example has 3 steps and 2nd fails
     Given an feature with one example with 3 steps where the 2nd step fails
     When the example is executed
     Then only two steps of the example were executed


  Rule: for given tags only examples with a matching tag are executed

  Example: feature with one matching and one non-matching example
    Given a feature has a first example "tagged example" with the tag "@sample"
    And the feature has a second example "un-tagged example" without any tag
    When the feature is executed with tag parameter "@sample"
    Then only the example "tagged example" is executed
