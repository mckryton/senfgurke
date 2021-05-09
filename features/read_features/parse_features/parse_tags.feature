Ability: parse tags
  Tags will help you to navigate through features and examples. For example it
  is possible to run only examples marked with a specific tag.
  Examples will inherit their tags from the feature. If a tag is set for a
  feature, all examples will have this tag too.

  Rule: tags preceeding an example should be assigned to this example

    Example: tags only on one example
      Given a feature
        """
          Feature: sample feature

            Example: first sample example

            @wip @important @beta
            Example: second sample example
        """
       When the feature is parsed
       Then the parsed features contains 2 examples
        And the parsed second example contains the tags "@wip, @important, @beta"

    Example: multiple tags above an example with a comment line inbetween
      Given a feature
        """
          Feature: sample feature

            @wip @important @beta
            # this is a comment
            Example: sample example
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
        """
      When the feature is parsed
      Then the parsed features contains an example
      And the parsed example contains the tags "@wip, @important, @beta"


  Rule: examples should inherit tags from the enclosing section
    Any tag preceding a feature or an example should be assigned accordingly
    any line starting with an @ sign is a tag line where tags are starting
    with @ and ending with space.

    Example: add inherited tags from a feature to an untagged example
      Given a feature
       """
          @wip @important @beta
          Feature: sample feature
            Example: an untagged example
       """
      When the feature is parsed
      Then the parsed features contains the tags "@wip, @important, @beta"
       And the included example contains the tags "@wip, @important, @beta"

    Example: add inherited tags from a rule to an untagged example
      Given a feature
       """
          Feature: sample feature

            @wip @important @beta
            Rule: a rule with tags

              Example: an untagged example
       """
      When the feature is parsed
      Then the parsed rule contains the tags "@wip, @important, @beta"
       And the included example contains the tags "@wip, @important, @beta"

    Example: add inherited tags from a feature and a rule to an untagged example
      Given a feature
       """
          Feature: sample feature

            @wip @important @beta
            Rule: a rule with tags

              Example: an untagged example
       """
      When the feature is parsed
      Then the parsed rule contains the tags "@wip, @important, @beta"
       And the included example contains the tags "@wip, @important, @beta"
