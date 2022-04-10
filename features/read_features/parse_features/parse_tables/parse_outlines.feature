@concept_outline
Ability: parse outlines
  Sometimes examples differ only in a small set of parameters. So instead of
  duplicating those examples for each set of parameter it is possible to add a
  table providing a set of parameters with each row at the end of a single
  example.


  Rule: an outline table is marked with the keyword "Examples:"

    @happy_path
    Example: outline table starting with an anonymous Examples keyword
      Given an example in a feature file is defined as
        """
          Example: climate scenario
            Given weather in <location> has <temperature> C

          Examples:
            | location | temperature |
            | Madrid   | 35          |
            | Munich   | 27          |
        """
       When the outline example is parsed
       Then the outline table is assigned to the example
        And the outline table has 2 columns named "location" and "temperature"
        And the data of row 1 of the outline table is "Madrid,35"
        And the data of row 2 of the outline table is "Munich,27"
        And the full name of the outline table is "Examples:"

    Example: outline table starting with a named Examples keyword
     Given an example in a feature file is defined as
       """
         Example: climate scenario
           Given weather in <location> has <temperature> C

         Examples: weather in Europe
           | location | temperature |
           | Madrid   | 35          |
           | Munich   | 27          |
       """
      When the outline example is parsed
      Then the outline table is assigned to the example
       And the full name of the outline table is "Examples: weather in Europe"

    Example: multiple anonymous outline tables
      Given an example in a feature file is defined as
        """
          Example: climate scenario
            Given weather in <location> has <temperature> C

          Examples:
            | location | temperature |
            | Madrid   | 35          |
            | Munich   | 27          |

          Examples:
            | location | temperature |
            | Cairo    | 42          |
            | Lagos    | 37          |
        """
       When the outline example is parsed
       Then 2 outline tables are assigned to the example
        And the full name of outline table 1 is "Examples:"
        And the full name of outline table 2 is "Examples:"

    Example: multiple named outline tables
      Given an example in a feature file is defined as
        """
          Example: climate scenario
            Given weather in <location> has <temperature> C

          Examples: weather in Europe
            | location | temperature |
            | Madrid   | 35          |
            | Munich   | 27          |

          Examples: weather in Africa
            | location | temperature |
            | Cairo    | 42          |
            | Lagos    | 37          |
        """
       When the outline example is parsed
       Then 2 outline tables are assigned to the example
        And the full name of outline table 1 is "Examples: weather in Europe"
        And the full name of outline table 2 is "Examples: weather in Africa"
