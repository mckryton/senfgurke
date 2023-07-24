Ability: connect steps with step functions
  Whenever a new step is added to a feature, Senfgurke will offer a matching
  step function. Senfgurke tries to make the function name as similar to the
  step name as possible so that it's easy to identify the matching step
  implementation for any step. But function names in any programming language
  have to follow conventions. Conventions for function and variable names in
  VBA are documented here:
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
  Rule: step function names should not exceed the max possible length set by VBA
    VBA spec says function names can have up to 255 characters
    (https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/visual-basic-naming-rules)
    but on MacOS VBA crashes for much shorter function names (eg. 59 characters
    on Excel 16.75.2)

   Example: step name not be shortened
     Given the max length for step function names is set to 58
       And a step "Given a step"
      When the step is translated into a function name
      Then the name of the resulting function is "Given_a_step_C72276450E70"

   Example: long step name to be shortened
     Given the max length for step function names is set to 50
       And a step "Given a very long step that is much much longer than the max length of step function names"
      When the step is translated into a function name
      Then the name of the resulting function is "Given_a_very_long_step_that_is_much_m_61CDF372E459"

  @vba-specific
  Rule: step function names replaces not allowed characters from VBA spec with underscores
    Note VBA on MacOS doesn't support UTF. For example the UTF-8 char "ü" (in
    hex = c3 bc or dec = 195 188) is converted into "√º" (chr(195) & chr(188)).

    Example: basic step with white spaces
      Given a step "Given a simple step"
       When the function for the step is calculated
       Then the function name starts with "Given_a_simple_step"

    Example: step uses umlauts
      Given a step "Given this step is überwältigend"
       When the function for the step is calculated
       Then the function name starts with "Given_this_step_is_berwltigend"

  @vba-specific
  Rule: step function names are made unique by adding a hashed value of the original step name
    Because the length of a function name is limited, function names derived
    from long steps are getting truncated. To keep the mapping between steps and
    step functions unique, step function names will include a hash value
    calculated from the whole step.

    Example: single letter step
      Given a step "Given x"
       When the function for the step is calculated
       Then the matching function name is "Given_x_6EF2FBE7CF29"

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
      Given a step "Given today is \"Monday\""
       When the function for the step is calculated
       Then the function name starts with "Given_today_is_STR_6A35E0F3D151"

    Example: step with escaped text
      Given a step
        """
          Given the next day is \"Monday\"
        """
       When the function for the step is calculated
       Then the function name starts with "Given_the_next_day_is_monday_"


  Rule: Gherkin steps should be case insensitive
    This should prevent having two steps that are different only in case get
    linked to two separate step functions.

   Example: Step descriptions differ only in case
     Given a list of steps
        | example_step     |
        | Given this is Up |
        | Given This is up |
        | Given This is UP |
      When the a matching function for those steps is requested
      Then the same function name is returned for all steps
