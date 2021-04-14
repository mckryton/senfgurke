@vba-specific
Ability: check if objects exists
    VBA distinguishes beteween variables assigned to objects o variables assigned
    to basic data types (e.g. string, intger, boolean). This expectation is
    validates if avariable is assigned to an object or not.

  Rule: is nothing expectation should fail for variables assigned to an object

    Example: expected nothing when object exists
      Given a variable refers to an object
      When the variable is expected to be nothing
      Then the expectation fails

    Example: expected nothing when object doesn't exists
      Given a variable refers to nothing
      When the variable is expected to be nothing
      Then the expectation is confirmed

    Example: expected nothing when value isn't an object
      Given a variable refers to 42
      When the non-object variable is expected to be nothing
      Then the expectation fails


  Rule: not is nothing expectation should confirm when a variable is assigned to an object

    Example: expected not nothing when object exists
      Given a variable refers to an object
       When the variable is expected not to be nothing
       Then the expectation is confirmed

    Example: expected not nothing when object doesn't exists
      Given a variable refers to nothing
       When the variable is expected not to be nothing
       Then the expectation fails

    Example: expected nothing when value isn't an object
      Given a variable refers to 42
       When the non-object variable is expected not to be nothing
       Then the expectation fails
