Ability: parse tags
  Senfgurke will read feature specs from text
  and identify tags assigned to the whole features or just examples

  #Todo: assign tags to rules
  Rule: any tag preceding a feature or an example should be assigned accordingly
    any line starting with an @ sign is a tag line where tags are starting
    with @ and ending with space

    Example: tags on the first example
      Given a feature
        """
          Feature: sample feature

            @wip @important @beta
            Example: sample example
        """
      When the feature is parsed
      Then the parsed features contains an example
      And the parsed example contains the tags "@wip, @important, @beta"


    Example: feature tags
      Given a feature
       """
          @wip @important @beta
          Feature: sample feature
       """
      When the feature is parsed
      Then the parsed features contains the tags "@wip, @important, @beta"


    Example: tags on the second example
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
