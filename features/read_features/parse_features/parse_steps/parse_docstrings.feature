Ability: parse docstrings
  Sometimes steps refer to a larger text strucure having multiple lines (e.g.
  describing a feature in a single step). Therefor step can be expanded by one
  docstring following the step directly. Docstrings are embraced by three double
  quotation marks


  Rule: add docstrings to the previous step
    A docstring is a multiline string that is related to the previous step.
    Docstrings are embraced by a sequence of 3 double quotation marks """

      Example: example with docstring
        Given an example
          And the first step is "Given a first step"
          And this step is followed by a docstring containing "this is a docstring"
         When the example is parsed
         Then first step is changed to "Given a first step \"this is a docstring\""

      Example: example with multiline docstring
        Given an example
              """
              Example: simple docstring
                Given a first step
              """
          And this is followed by a docstring
              """
                first line of docstring
                second line of docstring
              """
         When the example is parsed
         Then the step has one string parameter with two lines


      Example: feature background with docstring
        Given a background
          And the first step of the background is "Given a first background step"
          And this step is followed by a docstring containing "this is a docstring"
         When the feature background is parsed
         Then first step of the background is changed to "Given a first background step \"this is a docstring\""
