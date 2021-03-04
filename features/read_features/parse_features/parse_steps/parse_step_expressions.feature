Ability: parse step expressions
  To match a step with it's step definition it has to be disassembled into its
  parts: type (Given, When or Then) name (everything but the type) including
  any parameter values (step expressions) and an optional following docstring
  (a multiline text embraced by 3 double quotation marks in a row).


  Rule: text enclosed by double quotation marks should be a string expression

    Example: string expression
      Given a step definition "Given a step with a \"text parameter\""
       When the step definition is parsed
       Then the step has one step expression with the data type "string"
        And the value of the expression is "text parameter"

  Rule: any number that don't follows a character should be an long expression
    #Todo: Given step(2)
    #Todo: Given step[2]
    #Todo: Given step{2}
    #Todo: Given step #2

    Example: integer expression after space character
      Given a step definition "Given a step with a number 1"
       When the step definition is parsed
       Then the step has one step expression with the data type "long"
        And the value of the expression is 1

    Example: integer expression after brace character
      Given a step definition "Given the last paragraph is (2)"
       When the step definition is parsed
       Then the step has one step expression with the data type "long"
        And the value of the expression is 2

    Example: integer expressions in a list
      Given a step definition "Given a list contains (2,3)"
       When the step definition is parsed
       Then the step has 2 step expressions
        And the data types of the expressions are "long, long"
        And the value of the expressions are "2,3"

  Rule: any number that matches the int rule and encloses a single dot should be a double expression
    #Todo: Given step .34
    #Todo: add examples with braces and brackets

    Example: double expression
      Given a step definition "Given a step with the value of pi 3.14"
       When the step definition is parsed
       Then the step has one step expression with the data type "double"
        And the value of the expression is 3.14

    Example: step with a date
      Given a step definition "Given 01.01.2021 as a date"
       When the step definition is parsed
       Then the step has no step expressions
