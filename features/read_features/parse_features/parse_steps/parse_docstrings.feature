Ability: parse docstrings
  Sometimes steps refer to a larger text strucure having multiple lines (e.g.
  describing a feature in a single step). Therefor step can be expanded by one
  docstring following the step directly. Docstrings are embraced by three double
  quotation marks


  Rule: add docstrings to the previous step
    A docstring is a multiline string that is related to the previous step.
    Docstrings are embraced by a sequence of 3 double quotation marks

      Example: example with docstring
        Given an example definition
          And the first step is "Given a first step"
          And the step is followed by a docstring "this is a docstring"
         When the example is parsed
         Then the first step has an expression
          And the function name for the first step ends with "STR"

      Example: example with multiline docstring
        Given an example definition
              """
                Example: simple docstring
                  Given a first step
              """
          And the step is followed by a docstring
              """
                first line of docstring
                second line of docstring
              """
         When the example is parsed
         Then the step has one string parameter with two lines

      Example: docstring with leading linebreak
        Given a step definition "Given a sample step"
          And the step is followed by a docstring with a leading linebreak
         When the step is parsed
         Then line 1 of the steps docstring is empty

      Example: feature background with docstring
        Given a background
          And the first step is "Given a first background step"
          And the step is followed by a docstring "this is a docstring"
         When the feature background is parsed
         Then the first step has an expression
          And the function name for the first step ends with "STR"
          And the first step has a docstring value "this is a docstring"
