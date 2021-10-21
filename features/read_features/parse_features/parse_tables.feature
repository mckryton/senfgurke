Ability: parse tables
  Tables are a good way to keep examples concise and compact. Tables can be used
  as data table to extend steps or as outlines to extend examples.


  Rule: table rows are enclosed by a "|" while each column is separated by another "|"

    Example: table row aligned to the left
      Given a line "| sample row |" in a feature file
       When the type of this line is parsed
       Then the line is recognized as part of a table
        And the table row has one column

    Example: table row has only starting "|"
      Given a line "| sample row " in a feature file
       When the type of this line is parsed
       Then parsing will return the syntax error "closing | is missing in table row >| sample row<"

    @happy_path
    Example: table row with trailing white space
      Given a line "  | sample row | " in a feature file
       When the type of this line is parsed
       Then the line is recognized as part of a table
        And the table row has one column

    Example: table row with multiple columns
      Given a line "  | one | two | " in a feature file
       When the type of this line is parsed
       Then the line is recognized as part of a table
        And the table row has 2 columns

    Example: text before the first pipe
      Given a line "some text | sample row | " in a feature file
       When the type of this line is parsed
       Then the line is recognized as description


  Rule: the first line of a table contains column names with embedded white space

    @happy_path
    Example: table with two unique column names
      Given a step in a feature file is defined as
        """
          Given some weather is defined as
            | location | temperature in C |
        """
       When the step list is parsed
       Then the data table is assigned to the step
        And the table row has 2 columns
        And the names of the columns are "location" and "temperature in C"

    Example: table with duplicate column names
      Given a step in a feature file is defined as
        """
          Given some weather is defined as
            | location | location |
        """
       When the step list is parsed
       Then parsing will return the syntax error "duplicate column names found in data table: >location<"
