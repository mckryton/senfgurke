@vba-specific
Ability: check if objects exists
    VBA distinguishes between variables assigned to objects and variables
    assigned to basic data types (e.g. string, integer, boolean). This
    expectation validates if a variable is assigned to an object or not. The
    expectation is using the Nothing keyword to determin if a variable is
    assigned to an objoct or not.


  Rule: is Nothing expectation should fail for variables assigned to an object

    Example: expected Nothing when object exists
      Given a variable refers to an object
      When the variable is expected to be Nothing
      Then the expectation fails

    Example: expected Nothing when object doesn't exists
      Given a variable refers to Nothing
      When the variable is expected to be Nothing
      Then the expectation is confirmed

    Example: expected Nothing when value isn't an object
      Given a variable refers to 42
      When the non-object variable is expected to be Nothing
      Then the expectation fails


  Rule: not is Nothing expectation should confirm when a variable is assigned to an object

    Example: expected not Nothing when object exists
      Given a variable refers to an object
       When the variable is expected not to be Nothing
       Then the expectation is confirmed

    Example: expected not Nothing when object doesn't exists
      Given a variable refers to Nothing
       When the variable is expected not to be Nothing
       Then the expectation fails

    Example: expected Nothing when value isn't an object
      Given a variable refers to 42
       When the non-object variable is expected not to be Nothing
       Then the expectation fails
