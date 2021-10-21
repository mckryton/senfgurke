Ability: parse tags
  Tags help you navigate through features and examples. For example it is
  possible to run only examples marked with a specific tag.
  Examples inherit all tags from their features. If a tag is set for a
  feature, all examples in that feature will have this tag too.
  Any line starting with an @ sign is a tag line where tags are starting
  with @ and ending with space or linebreak.

  Rule: tags preceeding an example should be assigned to this example

    Example: tags only on one example
      Given a feature
        """
          Feature: sample feature

            Example: first sample example
              Given one step

            @wip @important @beta
            Example: second sample example
              Given one step
        """
       When the feature is parsed
       Then the parsed features contains 2 examples
        And the parsed second example is tagged with "@wip, @important, @beta"

    Example: multiple tags above an example with a comment line inbetween
      Given a feature
        """
          Feature: sample feature

            @wip @important @beta
            # this is a comment
            Example: sample example
              Given one Step
        """
       When the feature is parsed
       Then the parsed features contains an example
        And the parsed example contains the tags "@wip, @important, @beta"

    Example: multiple tags above an example with an empty line inbetween
      Given a feature
        """
          Feature: sample feature

            @wip @important @beta

            Example: sample example
              Given one step
        """
      When the feature is parsed
      Then the parsed features contains an example
      And the parsed example contains the tags "@wip, @important, @beta"


  Rule: examples should inherit tags from the enclosing section
    Any tag preceding a feature or an example should be assigned accordingly.

    Example: add inherited tags from a feature to an untagged example
      Given a feature
       """
          @wip @important @beta
          Feature: sample feature
            Example: an untagged example
              Given one step
       """
      When the feature is parsed
      Then the parsed features contains the tags "@wip, @important, @beta"
       And the included example contains the tags "@wip, @important, @beta"

    Example: add inherited tags from a rule to an untagged example for a feature with a single rule
      Given a feature
       """
          Feature: sample feature

            @wip @important @beta
            Rule: a rule with tags

              Example: an untagged example
                Given one step
       """
      When the feature is parsed
      Then the parsed rule contains the tags "@wip, @important, @beta"
       And the included example contains the tags "@wip, @important, @beta"

    Example: add inherited tags from a rule to an untagged example for a feature with multiple rules
     Given a feature
      """
         Feature: sample feature

           Rule: dummy rule
            Example: dummy example

           @wip @important @beta
           Rule: rule with tags
             Example: untagged example
               Given one step
      """
     When the feature is parsed
     Then the rule "rule with tags" contains the tags "@wip, @important, @beta"
      And the example "untagged example" contains the tags "@wip, @important, @beta" too

    Example: add inherited tags from a feature and a rule to an untagged example
      Given a feature
       """
          @alpha
          Feature: sample feature

            @wip @important @beta
            Rule: a rule with tags

              Example: an untagged example
                Given one step
       """
      When the feature is parsed
      Then the parsed rule contains the tags "@alpha, @wip, @important, @beta"
       And the included example contains the tags "@alpha, @wip, @important, @beta"
