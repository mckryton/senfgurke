Ability: parse tags
  Senfgurke will read feature specs from text
  and identify tag assigned to the whole features or just examples


  Rule: any line starting with an @ sign is a tag line where tags are starting with @ and ending with space

  Example: feature tags
    Given a feature
    And the first line of the feature is "  @wip @important @beta"
    When the feature is parsed
    Then the parsed features contains the tags wip, important and beta

  Example: example tags
    Given a feature
    And the line before the only example is "  @wip @important @beta"
    When the feature is parsed
    Then the parsed features contains an example
    And the example contains the tags wip, important and beta
