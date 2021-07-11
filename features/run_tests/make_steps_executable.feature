Ability: make steps executable
    Whenever a new step is added to a feature, Senfgurke will offer a matching
    step implementation function. Senfgurke tries to make the function name as
    similar to the step name as possible so that it's easy to identify the
    matching step implemenation for any step.
    But function names in any programming language have to follow conventions.
    Conventions for function and variable names in VBA are documented here:
    https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/visual-basic-naming-rules


  Rule: step function names should replace any "And" and "But" with a previous "Given", "When" or "Then"
    While looking at any step implementation function it should be obvious if
    it's a pre-condition, an action or a result.

    Example: "And" as synonym for "Given"
      Given an example
        """
          Example: sample
            Given a step
              And another step
        """
      When the function name for the last step is calculated
      Then the function name starts with "Given_another_step"

    Example: "But" as synonym for "Then"
      Given an example
        """
          Example: sample
             When the red light is turned on
              But no sound occurs
        """
       When the function name for the last step is calculated
       Then the function name starts with "When_no_sound_occurs"

  @vba-specific
  Rule: step function names replaces not allowed characters from VBA spec with underscores

    Example: basic step with white spaces

    Example: step uses special chars and umlauts

  @vba-specific
  Rule: step function names end with a hashed value of the original step name

  @vba-specific
  Rule: step expressions should appear as 3 uppercase letter abbreviation in function names
    Number and text values are used as parameters (expressions) for step
    function names.

    Example: Step with one integer value
      Given a step "Given a year has 12 months"
      When the function for the step is calculated
      Then the function name starts with "Given_a_year_has_INT_months"

    Example: Step with one floating value
      Given a step "Given the value of pi is 3.14"
      When the function for the step is calculated
      Then the function name starts with "Given_the_value_of_pi_is_DBL"

    Example: Step with one text value
      Given a step "Given the name of the first day of the week is \"Monday\""
      When the function for the step is calculated
      Then the function name starts with "Given_the_name_of_the_first_day_of_the_week_is_STR"

    Example: step with escaped text
