Ability: parse step expressions
  To match a step with it's step definition it has to be disassembled into its
  parts: type (Given, When or Then) name (everything but the type) including
  any parameter values (step expressions) and an optional following docstring
  (a multiline text embraced by """).


  Rule: recognice strings and numbers as step expressions

    Example: string expression
      Given a step definition "Given a step with a \"text parameter\""
       When the step definition is parsed
       Then the step has one step expression type "string"
        And the value of the expression is "text parameter"

    Example: integer expression
      Given a step definition "Given a step with a number 1"
       When the step definition is parsed
       Then the step has one step expression type "long"
        And the value of the expression is 1

    Example: double expression
      Given a step definition "Given a step with the value of pi 3.14"
       When the step definition is parsed
       Then the step has one step expression type "double"
        And the value of the expression is 3.14

    Example: step with a date
      Given a step definition "Given 01.01.2021 as a date"
       When the step definition is parsed
       Then the step has no step expressions
