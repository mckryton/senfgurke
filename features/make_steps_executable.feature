Ability: make steps executable
    Senfgurke will derive step function names from example steps so that a
    developer can add test code for each step.


  Rule: step function names don't use synonyms like and and but

    Example: And as synonym for Given
      Given a step "And another step"
      And the type of the previous step is "Given"
      When the function for the step is calculated
      Then the function name starts with "Given_another_step"

    Example: But as synonym for Then
      Given a step "But no sound occurs"
      And the type of the previous step is "When"
      When the function for the step is calculated
      Then the function name starts with "When_no_sound_occurs"


  Rule: step function names replaces not allowed characters from VBA spec with undersigns

    Example: basic step with white spaces

    Example: step uses special chars and umlauts


  Rule: step function names end with a hashed value of the original step name


  Rule: number and text values are used as parameters (expressions) for step function names

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

    @wip
    Example: add +1 to sum
      Given a is 2
      And b is 3
      When sum+1 is applied to a and b
      Then the result is 5
