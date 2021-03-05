Ability: parse comments
  To explain the intention of a feature it is possible to add a comment at any
  position. Any line starting with a # is considered to be such a comment. 


  Rule: every line starting with "#" is considered to be comment and will be ignored

    Example: feature starting with a comment
      Given a feature
        """
          # this is a comment
          Feature: sample feature
        """
      When the feature is parsed
      Then the parsed result contains a feature with the name "sample feature"

    Example: example with comment between steps
      Given a feature
        """
          Feature: sample feature

            Example: sample example
              Given one step
              # this is a comment
                And another step
        """
      When the feature is parsed
      Then the parsed result contains an example with two steps

    # add examples for comment between description lines
