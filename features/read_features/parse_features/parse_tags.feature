Ability: parse tags
<<<<<<< HEAD
  Tags will help you to navigate through features and examples. For example it
=======
  Tag will help you to navigate through features and examples. For example it
>>>>>>> f1bc461887479fc1a5bb5532ddeb9f5611672a3f
  is possible to run only examples marked with a specific tag.
  Examples will inherit their tags from the feature. If a tag is set for a
  feature, all examples will have this tag too.

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
