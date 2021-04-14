@vba-specific
Ability: confirm collection members
    A collection in VBA is like a named array (see
    https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object
    for more information). This expectation validates if a collection contains
    certain members or not.

  Rule: contains expectation should confirm that an array is an item of a collection

    Example: collection contains an array of primitive values
      Given a collection has 2 members array(1,"a") and array(2,"b")
      And an expected value is an array(2,"b")
      When the expectation validates that the collection contains the expected value
      Then the expectation is confirmed
