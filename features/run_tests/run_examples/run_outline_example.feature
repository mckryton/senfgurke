@concept_outline @wip
Ability: run outline example
  Examples can include outline tables to allow the repetition of an example but
  with a different set of parameters. So the example will be executed once for
  every row of the outline table while the value from each column will replace
  the corresponding placeholder in the examples steps.


  Rule: repeat examples for every row in an outline table

    Example: repeat example two times
      Given an example for an outline
        """
          Scenario: simple outline
            When something happens

          Examples:
            | col_1   |
            | value 1 |
            | value 2 |
        """
       When the example is executed
       Then in total 2 steps where executed

    Example: outline without data
      Given an example for an outline
        """
          Scenario: simple outline
            When something happens

          Examples:
            | col_1   |
        """
       When the example is executed
       Then in total 0 steps where executed

    Example: example has two outline tables
      Given an example for an outline
        """
         Scenario: simple outline
           When something happens

         Examples: one
           | col_1   |
           | value 1 |

         Examples: two
           | col_1   |
           | value 2 |
        """
       When the example is executed
       Then in total 2 steps where executed


  Rule: replace placeholders in step definitions with corresponding values from outline columns

    Example: step definition with a placeholder
      Given an example for an outline
        """
          Scenario: simple outline
            Given a value is set to <col_1>

          Examples:
            | col_1       |
            | replacement |
        """
       When the example is executed
       Then the step definition has changed to "Given a value is set to replacement"

   Example: example with multiple placeholders
     Given an example for an outline
       """
         Scenario: simple outline
           Given a number <col_1>
            When a function named <col_2> is called
            Then the result should be <col_3>

         Examples:
           | col_1 | col_2 | col_3 |
           | one   | add_g | gone  |
       """
      When the example is executed
      Then step definitions should have changed to
        """
          Given a number one
          When a function named add_g is called
          Then the result should be gone
        """

# TODO
#  Rule: run outlines only for matching tags
