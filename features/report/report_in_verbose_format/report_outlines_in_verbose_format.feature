@concept_outline
Ability: Report outlines in verbose format
    Because outlines are just slight variations of an example reporting them for
    each variation could make the report quite messy. Therefor the report should
    quote the original definition and only name the placeholders to be replaced
    with the values from the outline tables.

  Background:
    Given the report format is "verbose"

  Rule: steps for outline examples should be reported without upfront status but with placeholders

    Example: outline example with multiple executions
      Given multiple examples defined as an outline
        """
          Example: sample outline
            Given a step with has a value <param_1>

            Examples:
              | param_1 |
              | one     |
              | two     |
              | three   |
        """
       When the results from the execution of the outline examples are reported
       Then the report output of the example starts with
        """
          Example: sample outline
                    Given a step with has a value <param_1>
        """

      # another example could be an outline example embedded between regular
      # examples

  Rule: status for steps in an outline example should be reported as dots following each outline table row

    Example: outline example with multiple executions
      Given multiple examples defined as an outline
        """
          Example: sample outline
            Given a step with has a value <param_1>
             When <param_2> action happens

            Examples:
              | param_1 | param_2 |
              | one     | first   |
              | two     | second  |
        """
       When the results from the execution of the outline examples are reported with status "OK"
       Then the resulting report output is
        """
          Example: sample outline
                    Given a step with has a value <param_1>
                    When <param_2> action happens

                    Examples:
                      | param_1 | param_2 |
                      | one     | first   |
                      ..
                      | two     | second  |
                      ..
        """
