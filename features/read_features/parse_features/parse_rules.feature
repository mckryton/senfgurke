Ability: parse rules
  Quite often a feature has more than one aspect. E.g. most feature will follow
  several business policies. A rule will help you to reflect those aspects in
  the feature description. A rule describes the logic or policy a feature should
  follow. All examples following a rule should explain this policy. 
  A "Rule:" clause is limited by the start of the next rule or by the end of the
  feature.


  Rule: the rule keyword should mark the beginning of a "Rule:" clause
    The clause ends with the next "Rule:" clause or with the end of the feature

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
