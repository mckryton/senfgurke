Ability: parse rules
  A rule describes the logic or policy a feature shold follow. All examples
  following a rule should explain this policy. A rule is limited by the start
  of the next rule or by the end of the feature.

  Rule: the rule keyword should mark the boundary of a rule

    Example: single rule
      Given a feature
        """
          Feature: sample feature
            Rule: sample rule
        """
      When the feature is parsed
      Then the feature contains 1 rule(s)

    Example: two rules
    Given a feature
      """
        Feature: sample feature

          Rule: rule one
           this is the first rule

            Example: sample one
              Given one step

          Rule: rule two
           this is the second rule

            Example: sample two
              Given another step
      """
    When the feature is parsed
    Then the feature contains 2 rule(s)
     And each rule has 1 example

  Rule: examples following a rule should be attached to this rule
    # this is not yet implemented - at the moment rules are just syntactic sugar
