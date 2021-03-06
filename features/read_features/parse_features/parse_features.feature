Ability: parse features
  A feature describes the functionality of a software that will solve some
  specific problem of it's user. It does so by giving examples of how the
  function works.
  For a better understanding a featured follows a nested structure. A feature
  can contain zero or more rules (see parse rules for more). While a rule can
  contain zero or more examples (see parse examples for more). Of course only
  examples can be executed. So a feature or or rule without any assigned
  example will have no further effect on execution.
  Simple features can contain just some few examples without any rule.


  Rule: a feature should start with a valid keyword
    Features are valid only if they start with a "Feature:" keyword or one of
    its synonyms. Preceding tags or comments don't break this rule.

    Example: feature spec without feature keyword
      Given a feature
        """
          this is just some text
          without a feature keyword
        """
      When the feature is parsed
      Then the error "Feature lacks feature keyword at the beginning" was raised

    Example: feature with tags
      Given a feature
        """
          @tag1 @tag2
          Feature: sample feature
        """
      When the feature is parsed
      Then parsing didn't cause any errors

    Example: feature with leading empty lines
      Given a feature
        """


          Feature: sample feature
        """
      When the feature is parsed
      Then parsing didn't cause any errors

    Example: feature with leading comment line
      Given a feature
        """
          # this is a comment
          Feature: sample feature
        """
      When the feature is parsed
      Then parsing didn't cause any errors


  Rule: leading and trailing whitespace in clause definitions should be ignored
    Feature clauses are limited by lines matching this format
    <optional whitespace><keyword>:<optional whitespace><name><optional whitespace>
    A name can be any number of space-separated words, up-to the first newline.

    Example: keywords with whitespace
      Given a feature
        """
          Feature: sample feature
            Rule: sample rule
              Example: sample example
        """
      When the feature is parsed
      Then the parsed result contains a separate item for each of the given elements


  Rule: keyword values are single lines
    The name of rule or a feature is set in the same line as the corresponding
    keyword. Every following line not starting with a valid <keyword>: is just
    description.

    Example: feature with 2 lines of description
      Given a feature
        """
          Feature: sample feature
            this is
            the feature description
        """
       When the feature is parsed
       Then the feature description is set to those two lines

    Example: rule with description
      Given a feature
        """
          Feature: sample feature
            Rule: sample rule
              description for the rule
        """
       When the feature is parsed
       Then the parsed feature contains a rule
        And the rules description is set to "description for the rule"

    #TODO: add example for a line starting with an invalid keyword


  Rule: steps following a "Background:" clause should be assigned to the current rule or feature
    A background clause summarizes repeating steps in all examples.

    Example: feature background with one "Given" step
      Given a feature
        """
          Feature: sample feature
            Background:
              Given one step
        """
      When the feature is parsed
      Then the Given step is assigned to the feature





  # Rule: translate synonyms Ability,Business Needs to Feature
