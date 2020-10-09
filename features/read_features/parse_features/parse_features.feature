Ability: parse features
  Senfgurke will read feature specs from text
  and identify it's elements like descriptions, rules and examples

  # Rule: translate synonyms Ability,Business Needs to Feature
  Rule: parse features spec only if it's starting with "Feature:" keyword or synonyms or has preceding tags

    Example: feature spec without feature keyword
      Given a feature "this is just some text <br> without a feature keyword"
      When the feature is parsed
      Then the parsing results in the error message "Feature lacks feature keyword at the beginning"

    Example: feature with tags
      Given a feature "@tag1 @tag2<br>Feature: sample feature"
      When the feature is parsed
      Then the parsed feature doesn't contain any error


  Rule: feature clauses are limited by lines matching this format <optional whitespace><keyword>:<optional whitespace><name><optional whitespace>any optional whitespace is ignored

    Example: feature contains one rule and one example
      Given a feature named "sample feature"
      And the feature includes a rule "this is a sample rule"
      And the feature includes an example "this is an example"
      When the feature is parsed
      Then the parsed result contains a separate item for each of the given elements

    Example: keywords with whitespace
      Given a feature starting with "  Feature: sample feature "
      And the feature continues with the line "    Rule: this is a rule"
      When the feature is parsed
      Then the parsed result contains a feature with the name "sample feature"
      And the  parsed result contains a rule with the name "this is a rule"


  Rule: every line after a keyword up to the next keyword line or until the first example step in a example clause is considered to be description except comment lines

    Example: feature with 2 lines of description
      Given a feature named "sample feature"
      And the line with the feature keyword is followed by two lines "  this is<br>  the description"
      When the feature is parsed
      Then the feature description is set to those two lines

    Example: rule with description
      Given a feature named "sample feature"
      And the feature includes a rule "this is a sample rule"
      And the rule is followed by a line "  this is a description"
      When the feature is parsed
      Then the parsed feature contains a rule
      And the rules description is set to "this is a description"
