Ability: parse feature definition
  To distinct feature file from other text files a feature files shows its name
  on the top. Only tags, comments and empty lines are allowd above the feature
  name. Any description has to put below the feature name.
  The feature name is marked with any of the keywords reserved by Gherkin for
  features like feature, Abilty or Business Needs and a colon.


  Rule: a feature should start with a valid keyword
    Features are valid only if they start with a "Feature:" keyword or one of
    its synonyms. Preceding tags or comments don't break this rule.

    Example: feature spec without feature keyword
      Given a feature
        """
          this is just some text
          without a feature keyword
        """
      When the feature definition is parsed
      Then the error "feature lacks Feature keyword at the beginning" was raised

    Example: feature with tags
      Given a feature
        """
          @tag1 @tag2
          Feature: sample feature
        """
      When the feature definition is parsed
      Then parsing didn't cause any errors

    Example: feature with leading empty lines
      Given a feature
        """


          Feature: sample feature
        """
      When the feature definition is parsed
      Then parsing didn't cause any errors

    Example: feature with leading comment line
      Given a feature
        """
          # this is a comment
          Feature: sample feature
        """
      When the feature definition is parsed
      Then parsing didn't cause any errors
