@vba-specific
Ability: assure a collection has specific members
    A collection in VBA is like a named array (see
    https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object
    for more information). This expectation validates if a collection contains
    certain members or not.

  Rule: contains expectation should confirm that an array is an item of a collection
    Note: VBA know fixed sized arrays as well. Those arrays can be added to
    collections

    Example: collection contains an array of primitive values
      Given a collection has 2 members array(1,"a") and array(2,"b")
        And an expected value is an array(2,"b")
       When the expectation validates that the collection contains the expected value
       Then the expectation is confirmed

    Example: collections contains arrays but not the expected one
      Given a collection has 2 members array(1,"a") and array(2,"b")
        And an expected value is an array(3,"c")
       When the expectation validates that the collection contains the expected value
       Then the expectation fails

    Example: collections doesn't contain any array
      Given a collection has 2 members "a" and "b"
        And an expected value is an array(1,"a")
       When the expectation validates that the collection contains the expected value
       Then the expectation fails
