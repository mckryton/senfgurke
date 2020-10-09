Ability: parse comments
  Senfgurke will read feature specs from text
  and identify comments


  Rule: every line starting with "#" is considered to be comment and will be ignored

    Example: feature starting with a comment
      Given a feature named "sample feature"
      And the first line of the feature is "# this is a comment"
      When the feature is parsed
      Then the parsed result contains a feature with the name "sample feature"

    Example: example with comment between steps
      Given a feature named "sample feature"
      And the feature has an example with two steps and a comment line between the steps
      When the feature is parsed
      Then the parsed result contains an example with two steps

    # add examples for comment between description lines
