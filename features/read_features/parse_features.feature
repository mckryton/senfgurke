Ability: parse features
  Senfgurke will read feature specs from text
   and identify it's elements like descriptions, rules and examples


  # Rule: parse features spec only if it's starting with "Feature: keyword or synonyms or has preceding tags

  Example: feature spec without feature keyword"
    Given a feature "this is just some text <br> without a feature keyword"
    When the feature is parsed
    Then the parsing results in the error message "Feature lacks feature keyword at the beginning"


  # Rule: translate synonyms "Ability,Business Needs" to "Feature"

  Example: feature starts with "Ability:"
  

  # Rule: ignore trailing white space


  # Rule: Steps not starting with Given, When, Then or And will produce a Gherkin syntax error

  Example: missing step type
    Given an example with a step "bla bla bla"
    When the step is read
    Then the step result is an Gherkin Syntax Error claiming "bla" is not a valid step type


  # Rule: Translate "And" and "But" Given, When or Then using the same keyword as the previous step

  Example: And as synonym for Given
    Given an example with two steps "Given x is 1" and "And y is 2"
    When the second step is read
    Then the step type is set to "Given"

  Example: But as synonym for When
    Given an example with two steps "When some action happens" and "But it doesn't matter"
    When the second step is read
    Then the step type is set to "When"
