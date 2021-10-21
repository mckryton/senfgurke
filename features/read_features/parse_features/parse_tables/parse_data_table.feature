Ability: parse data table
  If examples require more extensive pre-conditions, brevity of the Given steps
  can be improved by using a data table instead of a lot of single steps.


  Rule: a data table is assigned to the preceding step
    # this rule doesn't consider scenario outlines (yet)

    @happy_path
    Example: data table follows a step line
      Given a step in a feature file is defined as
        """
          Given some weather is defined as
            | location | temperature in C |
            | Madrid   | 35               |
            | Munich   | 27               |
        """
       When the step list is parsed
       Then the data table is assigned to the step

    Example: data table follows a step line with empty line inbetween
     Given a step in a feature file is defined as
       """
         Given some weather is defined as

           | location | temperature in C |
           | Madrid   | 35               |
           | Munich   | 27               |
       """
      When the step list is parsed
      Then the data table is assigned to the step

    Example: data table follows example definition
     Given an example
       """
         Example: evaluate weather data
           | location | temperature in C |
           | Madrid   | 35               |
           | Munich   | 27               |
       """
      When the example is parsed
      Then parsing will return the syntax error "preceding step for data table is missing"
