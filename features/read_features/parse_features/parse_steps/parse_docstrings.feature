Ability: parse docstrings
  Sometimes steps refer to a larger text strucure having multiple lines (e.g.
  describing a feature in a single step). Therefore steps can be expanded by one
  docstring following the step directly. Docstrings are embraced by three double
  quotation marks


  Rule: add docstrings to the previous step as a step expression
    A docstring is a multiline string that is related to the previous step.
    Docstrings are embraced by a sequence of 3 double quotation marks

      Example: step list with a single line docstring
        Given a list of steps
            """
              Given a first step
                \"\"\"
                  this is a docstring
                \"\"\"
              And another step
            """
         When the step list is parsed
         Then the first step has an expression
          And the function name for the first step ends with "STR"

      Example: step list with a multi line docstring
        Given a list of steps
            """
              Given a first step
                \"\"\"
                  first line of docstring
                  second line of docstring
                \"\"\"
               And another step
            """
         When the step list is parsed
         Then the first step has an expression
          And the function name for the first step ends with "STR"
          And the first step has one string parameter with two lines

      Example: docstring with leading linebreak
        Given a list of steps
            """
              Given a sample step
                \"\"\"

                  this is a docstring
                \"\"\"
            """
         When the step list is parsed
         Then the first line of the steps docstring is empty
