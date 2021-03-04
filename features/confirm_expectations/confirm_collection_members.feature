@vba-specific
Ability: confirm collection members
    Senfgurke offers expectation functions to validate if a collection contains
    certain members or not.

  Rule: contains expectation should confirm that an array is an item of a collection

    Example: collection contains an array of primitive values
      Given a collection has 2 members array(1,"a") and array(2,"b")
      And an expected value is an array(2,"b")
      When the expectation validates that the collection contains the expected value
      Then the expectation is confirmed
