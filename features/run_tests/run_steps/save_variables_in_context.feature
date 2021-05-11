Ability: save variables in context
    Typically when implementing step definition functions there is a need to
    access the values defined in Given steps in the step definition functions
    of the When and Then steps. But steps for a single example can have theire
    step definitions in different classes, for example if the same step appears
    in more than one example. So to share values between step definition
    functions Senfgurke provides a context object as parameter for each step
    definition function that keeps this variable during the execution of a
    single example.

  Rule: context should save values for all data types between steps

    Example: string value in context
       When "hello world" is added as "greetings" in context
       Then the context value "greetings" returns "hello world"

    Example: integer value in context
       When 42 is add as "the answer" in context
       Then the context value "the answer" returns 42

    Example: object value in context
       When a new collection is added as "new list" to the context
       Then the context value "new list" returns a collection object


  Rule: context should clear all values after executing an example
